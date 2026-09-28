import { expect, it } from "vitest";
import { clearAIRequestDiagnostics, readAIRequestDiagnostics, readAIToolDiagnostics, recordAIRequest, recordAITool } from "./ai-diagnostics.js";
import { requestAnthropic } from "./anthropic.js";
it("retains only numerical diagnostics and collects provider cache usage", async () => {
  clearAIRequestDiagnostics();
  await requestAnthropic("secret-key", "customer-model", "secret-system", [], [{ role: "user", content: "SELECT private_data FROM customer; https://private.example" }], {}, undefined, { fetchImpl: async () => new Response(JSON.stringify({ content: [{ type: "text", text: "private-results" }], stop_reason: "end_turn", usage: { input_tokens: 10, output_tokens: 5, cache_read_input_tokens: 90, cache_creation_input_tokens: 4, private: "secret" } })) });
  expect(readAIRequestDiagnostics()[0]).toMatchObject({ inputTokens: 10, outputTokens: 5, cacheReadTokens: 90, cacheWriteTokens: 4, outcome: "complete" });
  recordAIRequest({ inputTokens: NaN, outputTokens: 0, cacheReadTokens: -10, cacheWriteTokens: 0, elapsedMs: Infinity, outcome: "error", secret: "credential", path: "/Users/private" } as any);
  recordAITool("customer_table", 12, "error");
  const serialized = JSON.stringify([readAIRequestDiagnostics(), readAIToolDiagnostics()]);
  expect(serialized).not.toMatch(/secret|customer|SELECT|private|https|credential|Users/);
  expect(readAIRequestDiagnostics()[1].elapsedMs).toBe(0);
  expect(readAIToolDiagnostics()[0]).toEqual({ kind: "catalog", elapsedMs: 12, outcome: "error" });
  clearAIRequestDiagnostics(); expect(readAIToolDiagnostics()).toEqual([]);
});
it("records streamed cache and output totals without retaining streamed content", async () => {
  clearAIRequestDiagnostics();
  const events = [ {type:"message_start",message:{usage:{input_tokens:3,cache_read_input_tokens:27}}}, {type:"content_block_start",content_block:{type:"text",text:""}}, {type:"content_block_delta",delta:{type:"text_delta",text:"secret"}}, {type:"content_block_stop"}, {type:"message_delta",delta:{stop_reason:"end_turn"},usage:{output_tokens:9}}, {type:"message_stop"} ];
  await requestAnthropic("key", "model", "", [], [], {}, undefined, {fetchImpl: async()=>new Response(events.map(event=>`data: ${JSON.stringify(event)}\n\n`).join(""),{headers:{"content-type":"text/event-stream"}})});
  expect(readAIRequestDiagnostics()[0]).toMatchObject({inputTokens:3,outputTokens:9,cacheReadTokens:27});
  expect(JSON.stringify(readAIRequestDiagnostics())).not.toContain("secret");
});
