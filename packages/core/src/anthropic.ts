import { aiUsage, recordAIRequest, type AIUsage } from "./ai-diagnostics.js";
import type { AIModelInfo } from "./ai-models.js";
/** Messages API behavior ported from vgi-web-frontend ac04853 / 1a251e1 / 5974f6d.
 * This module deliberately has no telemetry or credential persistence.
 */
export type AIEffort = "low" | "medium" | "high" | "xhigh" | "max";
export const AI_EFFORT_LEVELS: readonly AIEffort[] = ["low", "medium", "high", "xhigh", "max"];
export const DEFAULT_AI_EFFORT: AIEffort = "high";
export const DEFAULT_AI_MODEL = "claude-sonnet-5";
export const AI_MODELS = ["claude-haiku-4-5-20251001", "claude-sonnet-5", "claude-opus-5"] as const;
export function migrateAIModel(model: string): string {
  return model || DEFAULT_AI_MODEL;
}
export function normalizeEffort(value: unknown): AIEffort { return AI_EFFORT_LEVELS.includes(value as AIEffort) ? value as AIEffort : DEFAULT_AI_EFFORT; }
export function modelEfforts(model: string, info?: AIModelInfo): AIEffort[] {
  if (info?.id === model && info.efforts) return AI_EFFORT_LEVELS.filter(level => info.efforts!.includes(level));
  return model === "claude-sonnet-5" || model === "claude-opus-5" ? [...AI_EFFORT_LEVELS] : [];
}
export function supportsEffort(model: string, info?: AIModelInfo): boolean { return modelEfforts(model, info).length > 0; }
export function clampMaxTokens(model: string, requested = 16_384, info?: AIModelInfo): number {
  const ceiling = (info?.id === model ? info.maxTokens : undefined) ?? (model === "claude-haiku-4-5-20251001" ? 64_000 : supportsEffort(model) || ["claude-sonnet-4-6", "claude-opus-4-8"].includes(model) ? 128_000 : 8_192);
  return Math.min(ceiling, Number.isFinite(requested) && requested >= 1 ? Math.floor(requested) : 16_384);
}
export type TextBlock = { type: "text"; text: string };
export type ThinkingBlock = { type: "thinking"; thinking: string; signature: string } | { type: "redacted_thinking"; data: string };
export type ToolUseBlock = { type: "tool_use"; id: string; name: string; input: Record<string, unknown> };
export type AssistantBlock = TextBlock | ThinkingBlock | ToolUseBlock;
export type ToolResultBlock = { type: "tool_result"; tool_use_id: string; content: string; is_error?: boolean };
export type AgentBlock = AssistantBlock | ToolResultBlock;
export type AgentMessage = { role: "user" | "assistant"; content: string | AgentBlock[] };
export interface AIRequestOptions { modelInfo?: AIModelInfo; workspaceId?: string; effort?: AIEffort; maxTokens?: number; fetchImpl?: typeof fetch; endpoint?: string }
export interface AIStreamCallbacks { onThinking?(text: string): void; onUsage?(usage: AIUsage): void; onText?(text: string): void; onTool?(name: string, state: "writing", detail?: string, id?: string): void; onRetry?(message: string | null): void }

/** Complete missing results by addition, never by editing signed assistant blocks. */
export function repairAgentHistory(messages: AgentMessage[]): void {
  while (messages[0]?.role === "assistant") messages.shift();
  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    if (message.role !== "assistant" || !Array.isArray(message.content)) continue;
    const calls = message.content.filter((b): b is ToolUseBlock => b.type === "tool_use");
    if (!calls.length) continue;
    const next = messages[i + 1];
    const nextBlocks = next?.role === "user" ? asBlocks(next.content) : [];
    const answered = new Set(nextBlocks.filter((b): b is ToolResultBlock => b.type === "tool_result").map(b => b.tool_use_id));
    const missing: ToolResultBlock[] = calls.filter(b => !answered.has(b.id)).map(b => ({ type: "tool_result", tool_use_id: b.id, content: "This tool call was interrupted. It has no available result; inspect the current state before retrying.", is_error: true }));
    if (missing.length) {
      if (next?.role === "user") next.content = [...nextBlocks.filter(b => b.type === "tool_result"), ...missing, ...nextBlocks.filter(b => b.type !== "tool_result")];
      else messages.splice(i + 1, 0, { role: "user", content: missing });
    }
  }
  for (let i = 1; i < messages.length; i++) {
    if (messages[i - 1].role !== messages[i].role) continue;
    messages[i - 1].content = [...asBlocks(messages[i - 1].content), ...asBlocks(messages[i].content)];
    messages.splice(i--, 1);
  }
}
function asBlocks(content: AgentMessage["content"]): AgentBlock[] { return typeof content === "string" ? [{ type: "text", text: content }] : content; }

export function anthropicRequestBody(model: string, system: string, tools: readonly object[], messages: AgentMessage[], options: AIRequestOptions = {}): object {
  return {
    model, max_tokens: clampMaxTokens(model, options.maxTokens, options.modelInfo), stream: true,
    cache_control: { type: "ephemeral" },
    system: [{ type: "text", text: system, cache_control: { type: "ephemeral" } }],
    tools: tools.map((tool, index) => index === tools.length - 1 ? { ...tool, cache_control: { type: "ephemeral" } } : tool),
    messages,
    ...((options.modelInfo?.id === model ? options.modelInfo.adaptive ?? supportsEffort(model) : supportsEffort(model)) ? { thinking: { type: "adaptive", display: "summarized" } } : {}),
    ...(supportsEffort(model, options.modelInfo) ? { output_config: { effort: modelEfforts(model, options.modelInfo).includes(normalizeEffort(options.effort)) ? normalizeEffort(options.effort) : modelEfforts(model, options.modelInfo).includes("high") ? "high" : modelEfforts(model, options.modelInfo)[0] } } : {}),
  };
}

export async function requestAnthropic(apiKey: string, model: string, system: string, tools: readonly object[], messages: AgentMessage[], callbacks: AIStreamCallbacks, signal?: AbortSignal, options: AIRequestOptions = {}): Promise<{ content: AssistantBlock[]; stop_reason: string }> {
  const started = performance.now();
  let firstTextMs: number | undefined, usage = aiUsage(), outcome: "complete" | "stopped" | "error" = "error";
  try {
    const value = await requestAnthropicImpl(apiKey, model, system, tools, messages, { ...callbacks,
      onText: text => { firstTextMs ??= performance.now() - started; callbacks.onText?.(text); },
      onUsage: value => { usage = value; callbacks.onUsage?.(value); },
    }, signal, options);
    outcome = "complete"; return value;
  } catch (error) { if (signal?.aborted || (error as Error).name === "AbortError") outcome = "stopped"; throw error; }
  finally { recordAIRequest({ ...usage, elapsedMs: performance.now() - started, firstTextMs, outcome }); }
}

async function requestAnthropicImpl(apiKey: string, model: string, system: string, tools: readonly object[], messages: AgentMessage[], callbacks: AIStreamCallbacks, signal?: AbortSignal, options: AIRequestOptions = {}): Promise<{ content: AssistantBlock[]; stop_reason: string }> {
  if (!apiKey.trim()) throw new Error("An Anthropic API key is required.");
  signal?.throwIfAborted();
  const workspaceId = options.workspaceId?.trim();
  const response = await fetchAnthropic(options.endpoint ?? "https://api.anthropic.com/v1/messages", {
    method: "POST", signal,
    headers: { "content-type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true", ...(workspaceId ? { "anthropic-workspace-id": workspaceId } : {}) },
    body: JSON.stringify(anthropicRequestBody(model, system, tools, messages, options)),
  }, callbacks, options.fetchImpl ?? fetch);
  if (!response.headers.get("content-type")?.includes("text/event-stream")) {
    const value = await response.json() as { content: AssistantBlock[]; stop_reason: string; usage?: Record<string, unknown> };
    callbacks.onUsage?.(aiUsage(value.usage));
    if (value.stop_reason === "max_tokens") throw new Error("The AI response reached its output limit. Increase Max output tokens in AI settings or ask a smaller question.");
    for (const block of value.content) { if (block.type === "text") callbacks.onText?.(block.text); else if (block.type === "thinking" && block.thinking) callbacks.onThinking?.(block.thinking); else if (block.type === "tool_use") callbacks.onTool?.(block.name, "writing", undefined, block.id); }
    return value;
  }
  if (!response.body) throw new Error("Anthropic returned an empty response stream.");
  const reader = response.body.getReader(), decoder = new TextDecoder();
  const content: AssistantBlock[] = [];
  let usage: Record<string, unknown> = {};
  let buffer = "", current: AssistantBlock | null = null, json = "", stopReason = "end_turn", stopped = false;
  function eventLine(line: string): void {
    if (!line.startsWith("data:")) return;
    const data = line.slice(5).trim(); if (!data || data === "[DONE]") return;
    const event = JSON.parse(data);
    if (event.type === "message_start") { usage = { ...usage, ...event.message?.usage }; callbacks.onUsage?.(aiUsage(usage)); }
    if (event.type === "message_delta") { usage = { ...usage, ...event.usage }; callbacks.onUsage?.(aiUsage(usage)); }
    if (event.type === "content_block_start") {
      const block = event.content_block;
      current = null;
      if (block.type === "text") current = { type: "text", text: block.text ?? "" };
      else if (block.type === "thinking") { current = { type: "thinking", thinking: block.thinking ?? "", signature: block.signature ?? "" }; if (block.thinking) callbacks.onThinking?.(block.thinking); }
      else if (block.type === "redacted_thinking") current = { type: "redacted_thinking", data: block.data };
      else if (block.type === "tool_use") { current = { type: "tool_use", id: block.id, name: block.name, input: block.input ?? {} }; json = ""; callbacks.onTool?.(block.name, "writing", undefined, block.id); }
    } else if (event.type === "content_block_delta" && current) {
      const delta = event.delta;
      if (delta.type === "text_delta" && current.type === "text") { current.text += delta.text; callbacks.onText?.(delta.text); }
      else if (delta.type === "thinking_delta" && current.type === "thinking") { current.thinking += delta.thinking; callbacks.onThinking?.(delta.thinking); }
      else if (delta.type === "signature_delta" && current.type === "thinking") current.signature += delta.signature;
      else if (delta.type === "input_json_delta" && current.type === "tool_use") json += delta.partial_json;
    } else if (event.type === "content_block_stop" && current) {
      if (current.type === "tool_use" && json) {
        try { const parsed = JSON.parse(json); if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error(); current.input = parsed; }
        catch { current.input = { __parseError: "Invalid tool JSON" }; }
      }
      content.push(current); current = null;
    } else if (event.type === "message_delta") stopReason = event.delta?.stop_reason ?? stopReason;
    else if (event.type === "message_stop") stopped = true;
    else if (event.type === "error") throw new Error(event.error?.message ?? "Anthropic streaming error.");
  }
  const abort = () => { void reader.cancel().catch(() => undefined); };
  signal?.addEventListener("abort", abort, { once: true });
  try {
    while (true) {
      signal?.throwIfAborted();
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n"); buffer = lines.pop() ?? "";
      for (const line of lines) eventLine(line);
    }
    signal?.throwIfAborted();
    buffer += decoder.decode(); if (buffer.trim()) eventLine(buffer);
    if (current || !content.length) throw new Error("Anthropic response was interrupted. Please retry.");
    // Some proxies omit message_stop; complete content blocks remain usable.
    if (stopReason === "max_tokens") throw new Error("The AI response reached its output limit. Increase Max output tokens in AI settings or ask a smaller question.");
    return { content, stop_reason: stopReason };
  } finally { signal?.removeEventListener("abort", abort); if (!stopped) await reader.cancel().catch(() => undefined); reader.releaseLock(); }
}

async function fetchAnthropic(url: string, init: RequestInit, callbacks: AIStreamCallbacks, fetchImpl: typeof fetch): Promise<Response> {
  try {
    for (let attempt = 0; ; attempt++) {
      init.signal?.throwIfAborted();
      let response: Response;
      try { response = await fetchImpl(url, init); }
      catch (error) {
        init.signal?.throwIfAborted();
        if (!(error instanceof TypeError) || attempt >= 3) throw error;
        const seconds = Math.min(10, 2 ** attempt);
        callbacks.onRetry?.(`Network interruption; retrying in ${seconds}s…`);
        await delay(seconds * 1000, init.signal); continue;
      }
      if (response.ok) return response;
      const body = await response.text();
      if (response.status === 400 && body.includes("anthropic-workspace-id is required")) throw new Error("This API key requires an Anthropic workspace ID. Add it in AI settings.");
      if (response.status === 400 && body.includes("anthropic-workspace-id")) throw new Error("Invalid Anthropic workspace ID. Check the wrkspc_… value in AI settings.");
      if (response.status === 404 && /workspace .*not found/i.test(body)) throw new Error("Anthropic workspace not found, or this API key does not have access to it. Check AI settings.");
      if (![429, 529].includes(response.status) || attempt >= 3) throw new Error(`Anthropic request failed (${response.status}): ${body}`);
      const seconds = Math.min(30, Math.max(1, Number(response.headers.get("retry-after")) || 2 ** attempt));
      callbacks.onRetry?.(`Anthropic is busy; retrying in ${seconds}s…`);
      await delay(seconds * 1000, init.signal);
    }
  } finally { callbacks.onRetry?.(null); }
}
function delay(ms: number, signal?: AbortSignal | null): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) { reject(signal.reason ?? new DOMException("Stopped", "AbortError")); return; }
    const abort = () => { clearTimeout(timer); reject(signal?.reason ?? new DOMException("Stopped", "AbortError")); };
    const timer = setTimeout(() => { signal?.removeEventListener("abort", abort); resolve(); }, ms);
    signal?.addEventListener("abort", abort, { once: true });
  });
}
