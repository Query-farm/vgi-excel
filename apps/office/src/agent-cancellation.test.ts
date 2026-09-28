import { expect, it, vi } from "vitest";
import type { QueryBackend, QueryResult } from "@query-farm/vgi-excel-core";
import { OfficeAgentSession } from "./anthropic";
import { agentContext } from "./catalog";

it.each(["inventory", "sql", "catalog"])("interrupts Office agent %s and permits a follow-up", async phase => {
  const abort = new AbortController();
  let started!: () => void;
  const waiting = new Promise<void>(resolve => { started = resolve; });
  let block = true;
  const signals: AbortSignal[] = [];
  const empty = { columns: [], rows: [], rowCount: 0 };
  const backend: QueryBackend = {
    call: async () => empty,
    query: vi.fn(async (sql, options) => {
      if (block && (phase === "inventory" || (phase === "sql" ? sql === "SELECT 42" : sql.includes("CAST(to_json(parameters)")))) {
        expect(options?.signal).toBe(abort.signal);
        signals.push(options!.signal!); started();
        return new Promise<QueryResult>((_, reject) => options!.signal!.addEventListener("abort", () => reject(new DOMException("Stopped", "AbortError")), { once: true }));
      }
      return empty;
    }),
  };
  const fetchImpl = vi.fn(async () => new Response(JSON.stringify({ content: [{ type: "tool_use", id: "cancel-tool", name: phase === "catalog" ? "list_functions" : "run_sql", input: { sql: "SELECT 42" } }], stop_reason: "tool_use" }), { headers: { "content-type": "application/json" } }));
  const session = new OfficeAgentSession(), onResult = vi.fn();
  const run = session.run("key", "Start", agentContext(backend), abort.signal, { fetchImpl }, { onResult });
  const rejected = expect(run).rejects.toMatchObject({ name: "AbortError" });
  await waiting; abort.abort(); await rejected;
  expect(signals.length).toBeGreaterThan(0);
  expect(onResult).not.toHaveBeenCalled();
  expect(fetchImpl).toHaveBeenCalledTimes(phase === "inventory" ? 0 : 1);
  block = false;
  const onText = vi.fn();
  const resumed = await session.run("key", "Continue", agentContext(backend), undefined, {
    fetchImpl: async () => new Response(JSON.stringify({ content: [{ type: "text", text: "Recovered" }], stop_reason: "end_turn" }), { headers: { "content-type": "application/json" } }),
  }, { onText });
  expect(resumed.text).toBe("Recovered");
  if (phase !== "inventory") expect(JSON.stringify(session.snapshot())).toContain('"is_error":true');
});
