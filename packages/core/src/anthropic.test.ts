import { describe, expect, it, vi } from "vitest";
import { anthropicRequestBody, clampMaxTokens, migrateAIModel, normalizeEffort, repairAgentHistory, requestAnthropic, type AgentMessage } from "./anthropic.js";
import { tableClipboard } from "./table-clipboard.js";

const messages: AgentMessage[] = [{ role: "user", content: "Hello" }];
function sse(events: object[]): Response {
  const text = events.map(event => `data: ${JSON.stringify(event)}\n\n`).join("");
  // Split inside JSON and signatures, rather than assuming network chunk boundaries.
  const bytes = new TextEncoder().encode(text);
  return new Response(new ReadableStream({ start(controller) { for (let i = 0; i < bytes.length; i += 7) controller.enqueue(bytes.slice(i, i + 7)); controller.close(); } }), { headers: { "content-type": "text/event-stream" } });
}
const thinkingEvents = [
  { type: "content_block_start", content_block: { type: "thinking", thinking: "", signature: "" } },
  { type: "content_block_delta", delta: { type: "thinking_delta", thinking: "Consider λ" } },
  { type: "content_block_delta", delta: { type: "signature_delta", signature: "first" } },
  { type: "content_block_delta", delta: { type: "signature_delta", signature: "-second" } },
  { type: "content_block_stop" },
  { type: "content_block_start", content_block: { type: "redacted_thinking", data: "opaque" } },
  { type: "content_block_stop" },
  { type: "content_block_start", content_block: { type: "tool_use", id: "call", name: "list_tables", input: {} } },
  { type: "content_block_stop" },
  { type: "message_delta", delta: { stop_reason: "tool_use" } },
  { type: "message_stop" },
];

describe("Anthropic upstream ports", () => {
  it("gates adaptive thinking and effort and clamps output limits", () => {
    expect(anthropicRequestBody("claude-sonnet-5", "stable", [], messages, { effort: "max" })).toMatchObject({ thinking: { type: "adaptive", display: "summarized" }, output_config: { effort: "max" }, max_tokens: 16384 });
    expect(anthropicRequestBody("claude-haiku-4-5-20251001", "stable", [], messages)).not.toHaveProperty("thinking");
    expect(anthropicRequestBody("custom-model", "stable", [], messages)).not.toHaveProperty("output_config");
    expect(clampMaxTokens("claude-haiku-4-5-20251001", 128000)).toBe(64000);
    expect(clampMaxTokens("custom-model", Infinity)).toBe(8192);
    expect(clampMaxTokens("claude-sonnet-5", 0.4)).toBe(16384);
    expect(normalizeEffort("invalid")).toBe("high");
    expect(migrateAIModel("claude-sonnet-4-6")).toBe("claude-sonnet-4-6");
    expect(migrateAIModel("custom-model")).toBe("custom-model");
  });
  it("streams only the provider thinking summary while preserving signed blocks", async () => {
    const onText = vi.fn(), onTool = vi.fn(), onThinking = vi.fn();
    const result = await requestAnthropic("key", "claude-sonnet-5", "stable", [], messages, { onText, onTool, onThinking }, undefined, { fetchImpl: async () => sse(thinkingEvents) });
    expect(result.content).toEqual([
      { type: "thinking", thinking: "Consider λ", signature: "first-second" },
      { type: "redacted_thinking", data: "opaque" },
      { type: "tool_use", id: "call", name: "list_tables", input: {} },
    ]);
    expect(onThinking.mock.calls).toEqual([["Consider λ"]]);
    expect(onText).not.toHaveBeenCalled();
    expect(onTool).toHaveBeenCalledWith("list_tables", "writing", undefined, "call");
  });
  it("adds interrupted tool results without rewriting signed assistant history", () => {
    const history: AgentMessage[] = [...messages, { role: "assistant", content: [{ type: "thinking", thinking: "", signature: "signed" }, { type: "tool_use", id: "a", name: "run_sql", input: { sql: "SELECT 1" } }, { type: "tool_use", id: "b", name: "run_sql", input: { sql: "SELECT 2" } }] }, { role: "user", content: [{ type: "tool_result", tool_use_id: "a", content: "already completed" }] }];
    const before = JSON.stringify(history[1]);
    repairAgentHistory(history);
    expect(JSON.stringify(history[1])).toBe(before);
    expect(history[2].content).toEqual([ { type: "tool_result", tool_use_id: "a", content: "already completed" }, expect.objectContaining({ type: "tool_result", tool_use_id: "b", is_error: true }) ]);
    const repaired = JSON.stringify(history); repairAgentHistory(history); expect(JSON.stringify(history)).toBe(repaired);
  });
  it("marks stable prefixes without mutating the tool list or messages", () => {
    const tools = [{ name: "a" }, { name: "b" }];
    const before = JSON.stringify({ tools, messages });
    const body = anthropicRequestBody("claude-sonnet-5", "stable", tools, messages) as any;
    expect(body.cache_control).toEqual({ type: "ephemeral" });
    expect(body.tools[1].cache_control).toEqual({ type: "ephemeral" });
    expect(body.system[0].cache_control).toEqual({ type: "ephemeral" });
    expect(JSON.stringify({ tools, messages })).toBe(before);
  });
  it.each([undefined, "  wrkspc_example  "])("sends workspace credentials only in headers (%s)", async workspaceId => {
    const fetchImpl = vi.fn(async () => sse(thinkingEvents));
    await requestAnthropic("secret-key", "claude-sonnet-5", "stable", [], messages, {}, undefined, { workspaceId, fetchImpl });
    const init = (fetchImpl.mock.calls as unknown as [string, RequestInit][])[0][1];
    expect(new Headers(init.headers).get("anthropic-workspace-id")).toBe(workspaceId ? "wrkspc_example" : null);
    expect(String(init.body)).not.toContain("secret-key");
    expect(String(init.body)).not.toContain("wrkspc_example");
  });
  it("gives workspace errors without retrying", async () => {
    const fetchImpl = vi.fn(async () => new Response('{"error":{"message":"anthropic-workspace-id is required"}}', { status: 400 }));
    await expect(requestAnthropic("key", "model", "stable", [], messages, {}, undefined, { fetchImpl })).rejects.toThrow("Add it in AI settings");
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
  it("cancels a retry wait immediately", async () => {
    const abort = new AbortController();
    const fetchImpl = vi.fn(async () => new Response("busy", { status: 429, headers: { "retry-after": "30" } }));
    await expect(requestAnthropic("key", "model", "stable", [], messages, { onRetry: status => { if (status) abort.abort(); } }, abort.signal, { fetchImpl })).rejects.toMatchObject({ name: "AbortError" });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
  it("reports token truncation rather than executing incomplete tools", async () => {
    const events = [...thinkingEvents.slice(0, -2), { type: "message_delta", delta: { stop_reason: "max_tokens" } }, { type: "message_stop" }];
    await expect(requestAnthropic("key", "model", "stable", [], messages, {}, undefined, { fetchImpl: async () => sse(events) })).rejects.toThrow("output limit");
  });
  it("copies table headers and safely escapes rich HTML and multiline TSV", () => {
    const result = tableClipboard([["city", "note"], ["<Paris>", 'first\n"second"']]);
    expect(result.text).toBe('city\tnote\n<Paris>\t"first\n""second"""');
    expect(result.html).toContain("&lt;Paris&gt;");
    expect(result.html).not.toContain("<Paris>");
  });
});

it("cancels a blocked SSE reader after partial text without retrying", async () => {
  const controller = new AbortController();
  const cancelled = vi.fn();
  const fetchImpl = vi.fn(async () => new Response(new ReadableStream({
    start(stream) { stream.enqueue(new TextEncoder().encode('data: {"type":"content_block_start","content_block":{"type":"text","text":""}}\n\ndata: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Partial"}}\n\n')); },
    cancel: cancelled,
  }), { headers: { "content-type": "text/event-stream" } }));
  const text = vi.fn(() => controller.abort());
  await expect(requestAnthropic("key", "model", "system", [], messages, { onText: text }, controller.signal, { fetchImpl })).rejects.toMatchObject({ name: "AbortError" });
  expect(text).toHaveBeenCalledWith("Partial");
  expect(cancelled).toHaveBeenCalledTimes(1);
  expect(fetchImpl).toHaveBeenCalledTimes(1);
});

it("keeps JSON thinking summaries separate from answer text and redacted data", async () => {
  const onThinking = vi.fn(), onText = vi.fn();
  await requestAnthropic("key", "claude-sonnet-5", "", [], messages, { onThinking, onText }, undefined, {
    fetchImpl: async () => new Response(JSON.stringify({ content: [
      { type: "thinking", thinking: "Checking the requested period.", signature: "private-signature" },
      { type: "redacted_thinking", data: "private-redacted" },
      { type: "text", text: "The result is ready." }
    ], stop_reason: "end_turn" }))
  });
  expect(onThinking.mock.calls).toEqual([["Checking the requested period."]]);
  expect(onText.mock.calls).toEqual([["The result is ready."]]);
});
