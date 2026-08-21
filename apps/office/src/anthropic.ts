import { AGENT_TOOLS, agentRoundSeparator, appendAgentRoundText, buildExcelAgentSystemPrompt, executeAgentTool, type AgentCatalogObject, type AgentConnectionContext, type AgentToolContext, type QueryResult } from "@query-farm/vgi-excel-core";

export type OfficeContentBlock =
  | { type: "text"; text: string }
  | { type: "tool_use"; id: string; name: string; input: Record<string, unknown> };
type ContentBlock = OfficeContentBlock;

interface AnthropicResponse {
  content: ContentBlock[];
  stop_reason: string;
}

export interface OfficeAgentMessage {
  role: "user" | "assistant";
  content: string | unknown[];
}
type Message = OfficeAgentMessage;

export interface AgentAnswer {
  text: string;
  stagedResult?: QueryResult;
}

export interface AgentOptions {
  endpoint?: string;
  model?: string;
  fetchImpl?: typeof fetch;
  connection?: AgentConnectionContext;
}

export interface AgentCallbacks {
  onText?(chunk: string): void;
  onTool?(name: string, state: "writing" | "running" | "done" | "error", detail?: string, id?: string): void;
  onResult?(result: QueryResult): void;
}

export class OfficeAgentSession {
  private readonly messages: Message[] = [];

  reset(): void { this.messages.length = 0; }
  restore(messages: OfficeAgentMessage[]): void { this.messages.splice(0, this.messages.length, ...JSON.parse(JSON.stringify(messages))); sanitizeMessages(this.messages); }
  snapshot(): OfficeAgentMessage[] { const value = JSON.parse(JSON.stringify(this.messages)) as OfficeAgentMessage[]; sanitizeMessages(value); return value; }

  async run(apiKey: string, prompt: string, context: AgentToolContext, signal?: AbortSignal, options: AgentOptions = {}, callbacks: AgentCallbacks = {}): Promise<AgentAnswer> {
    this.messages.push({ role: "user", content: prompt });
    sanitizeMessages(this.messages);
    const systemPrompt = await loadSystemPrompt(context, options.connection);
    let answer = "";
    let stagedResult: QueryResult | undefined;
    for (let round = 0; round < 12; round++) {
      let firstTextChunk = true;
      const roundCallbacks: AgentCallbacks = { ...callbacks, onText: (chunk) => {
        if (firstTextChunk && chunk) { firstTextChunk = false; const separator = agentRoundSeparator(answer, chunk); if (separator) callbacks.onText?.(separator); }
        callbacks.onText?.(chunk);
      } };
      const response = await callAnthropic(apiKey, this.messages, systemPrompt, signal, options, roundCallbacks);
      this.messages.push({ role: "assistant", content: response.content });
      const text = response.content.filter((block): block is Extract<ContentBlock, { type: "text" }> => block.type === "text").map((block) => block.text).join("");
      answer = appendAgentRoundText(answer, text);
      const calls = response.content.filter((block): block is Extract<ContentBlock, { type: "tool_use" }> => block.type === "tool_use");
      if (!calls.length) return { text: answer, stagedResult };
      const results: unknown[] = [];
      for (const call of calls) {
        callbacks.onTool?.(call.name, "running", summarizeInput(call.input), call.id);
        const started = performance.now();
        try {
          const result = await executeAgentTool(call.name, call.input, context);
          if (result.queryResult) { stagedResult = result.queryResult; callbacks.onResult?.(result.queryResult); }
          callbacks.onTool?.(call.name, "done", `${Math.round(performance.now() - started)} ms`, call.id);
          results.push({ type: "tool_result", tool_use_id: call.id, content: result.content });
        } catch (error) {
          callbacks.onTool?.(call.name, "error", message(error), call.id);
          results.push({ type: "tool_result", tool_use_id: call.id, is_error: true, content: message(error) });
        }
      }
      this.messages.push({ role: "user", content: results });
    }
    throw new Error("The agent exceeded the maximum number of tool rounds.");
  }
}

function sanitizeMessages(messages: OfficeAgentMessage[]): void {
  while (messages[0]?.role === "assistant") messages.shift();
  for (let index = 0; index < messages.length; index++) {
    const message = messages[index];
    if (message.role !== "assistant" || !Array.isArray(message.content)) continue;
    const calls = message.content.filter((block): block is { type: "tool_use"; id: string } => !!block && typeof block === "object" && (block as { type?: string }).type === "tool_use" && typeof (block as { id?: unknown }).id === "string");
    if (!calls.length) continue;
    const next = messages[index + 1];
    const answered = new Set(Array.isArray(next?.content) ? next.content.filter((block) => !!block && typeof block === "object" && (block as { type?: string }).type === "tool_result").map((block) => String((block as { tool_use_id?: unknown }).tool_use_id ?? "")) : []);
    message.content = message.content.filter((block) => !block || typeof block !== "object" || (block as { type?: string }).type !== "tool_use" || answered.has(String((block as { id?: unknown }).id ?? "")));
    if (!message.content.length) message.content = [{ type: "text", text: "(stopped)" }];
  }
  for (let index = 1; index < messages.length; index++) {
    const previous = messages[index - 1], current = messages[index];
    if (previous.role !== current.role) continue;
    previous.content = [...contentBlocks(previous.content), ...contentBlocks(current.content)];
    messages.splice(index, 1); index--;
  }
}

function contentBlocks(content: OfficeAgentMessage["content"]): unknown[] { return Array.isArray(content) ? content : content ? [{ type: "text", text: content }] : []; }

export async function runAgent(
  apiKey: string,
  prompt: string,
  context: AgentToolContext,
  signal?: AbortSignal,
  options: AgentOptions = {},
): Promise<AgentAnswer> {
  return new OfficeAgentSession().run(apiKey, prompt, context, signal, options);
}

async function callAnthropic(apiKey: string, messages: Message[], systemPrompt: string, signal: AbortSignal | undefined, options: AgentOptions, callbacks: AgentCallbacks): Promise<AnthropicResponse> {
  const response = await (options.fetchImpl ?? fetch)(options.endpoint ?? "https://api.anthropic.com/v1/messages", {
    method: "POST",
    signal,
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: options.model ?? "claude-sonnet-4-5",
      max_tokens: 8192,
      stream: true,
      system: systemPrompt,
      tools: AGENT_TOOLS,
      messages,
    }),
  });
  if (!response.ok) throw new Error(`Anthropic request failed (${response.status}): ${await response.text()}`);
  if (!response.headers.get("content-type")?.includes("text/event-stream")) {
    const value = await response.json() as AnthropicResponse;
    for (const block of value.content) { if (block.type === "text") callbacks.onText?.(block.text); else callbacks.onTool?.(block.name, "writing", undefined, block.id); }
    return value;
  }
  if (!response.body) throw new Error("Anthropic returned an empty response stream.");
  const blocks: ContentBlock[] = [];
  let current: ContentBlock | null = null, toolJson = "", buffer = "", stopReason = "end_turn";
  const reader = response.body.getReader(), decoder = new TextDecoder();
  while (true) {
    const { done, value } = await reader.read(); if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n"); buffer = lines.pop() ?? "";
    for (const line of lines) {
      if (!line.startsWith("data: ")) continue;
      const raw = line.slice(6).trim(); if (!raw || raw === "[DONE]") continue;
      const event = JSON.parse(raw) as any;
      if (event.type === "content_block_start") {
        if (event.content_block?.type === "text") current = { type: "text", text: "" };
        else if (event.content_block?.type === "tool_use") { current = { type: "tool_use", id: event.content_block.id, name: event.content_block.name, input: {} }; toolJson = ""; callbacks.onTool?.(current.name, "writing", undefined, current.id); }
      } else if (event.type === "content_block_delta" && current) {
        if (event.delta?.type === "text_delta" && current.type === "text") { current.text += event.delta.text; callbacks.onText?.(event.delta.text); }
        else if (event.delta?.type === "input_json_delta" && current.type === "tool_use") toolJson += event.delta.partial_json;
      } else if (event.type === "content_block_stop" && current) {
        if (current.type === "tool_use") { try { current.input = JSON.parse(toolJson || "{}"); } catch { current.input = {}; } }
        blocks.push(current); current = null;
      } else if (event.type === "message_delta") stopReason = event.delta?.stop_reason ?? stopReason;
      else if (event.type === "error") throw new Error(event.error?.message ?? "Anthropic streaming error.");
    }
  }
  return { content: blocks, stop_reason: stopReason };
}

async function loadSystemPrompt(context: AgentToolContext, connection?: AgentConnectionContext): Promise<string> {
  const active = connection ?? { name: "selected VGI connection", catalog: "selected catalog", authentication: "unspecified" };
  const objects: AgentCatalogObject[] = [];
  const errors: string[] = [];
  const [tables, functions] = await Promise.allSettled([context.listTables(), context.listFunctions()]);
  if (tables.status === "fulfilled") {
    for (const row of records(tables.value)) objects.push({
      catalog: String(row.table_catalog ?? active.catalog), schema: String(row.table_schema ?? "main"),
      name: String(row.table_name ?? ""), kind: String(row.table_type ?? "table").toLowerCase(),
    });
  } else errors.push(message(tables.reason));
  if (functions.status === "fulfilled") {
    for (const row of records(functions.value)) objects.push({
      catalog: String(row.catalog ?? active.catalog), schema: String(row.schema ?? "main"), name: String(row.name ?? ""),
      kind: String(row.kind ?? "function"), description: row.description == null ? undefined : String(row.description),
    });
  } else errors.push(message(functions.reason));
  return buildExcelAgentSystemPrompt({ connection: active, objects: objects.filter((item) => item.name), inventoryError: errors.join("; ") || undefined });
}

function records(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value) ? value.filter((item): item is Record<string, unknown> => !!item && typeof item === "object" && !Array.isArray(item)) : [];
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function summarizeInput(input: Record<string, unknown>): string {
  const value = JSON.stringify(input); return value.length > 500 ? value.slice(0, 500) + "…" : value;
}
