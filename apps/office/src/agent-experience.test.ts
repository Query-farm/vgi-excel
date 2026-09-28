import { expect, it, vi } from "vitest";
import { OfficeAgentSession } from "./anthropic";
import type { AgentToolContext } from "@query-farm/vgi-excel-core";
const empty = { columns: [], rows: [], rowCount: 0 };
function context(): AgentToolContext { return { backend: { query: vi.fn(async () => empty), call: async () => empty }, listTables: vi.fn(async () => []), listFunctions: async () => [], describeTable: async () => [] }; }
function tool(id: string, name: string, input: object) { return new Response(JSON.stringify({ content: [{ type: "tool_use", id, name, input }], stop_reason: "tool_use" })); }
const done = () => new Response(JSON.stringify({ content: [{ type: "text", text: "Done" }], stop_reason: "end_turn" }));
it("waits for clarification before dependent SQL and attaches the executed SQL and scope", async () => {
  const ctx = context(), session = new OfficeAgentSession(), onQueryResult = vi.fn();
  let answer!: (value: string) => void;
  const onClarification = vi.fn(() => new Promise<string>(resolve => { answer = resolve; }));
  const fetchImpl = vi.fn().mockImplementationOnce(() => tool("q", "ask_clarification", { question: "Which period?", options: ["This year", "Last year"] })).mockImplementationOnce(() => tool("sql", "run_sql", { sql: "SELECT 42", scope: { dateRange: "Last year" } })).mockImplementationOnce(done);
  const run = session.run("key", "Revenue", ctx, undefined, { fetchImpl, connection: { name: "source", catalog: "catalog" } }, { onClarification, onQueryResult });
  await vi.waitFor(() => expect(onClarification).toHaveBeenCalledTimes(1));
  expect(ctx.backend.query).not.toHaveBeenCalled(); expect(fetchImpl).toHaveBeenCalledTimes(1);
  answer("Last year"); await run;
  expect(onQueryResult).toHaveBeenCalledWith(expect.objectContaining({ sql: "SELECT 42", connection: "source", scope: { dateRange: "Last year" } }));
  expect(JSON.stringify(session.snapshot())).toContain("Last year");
});
it("bounds repeated tools and resets the guard for a new turn", async () => {
  const ctx = context(), session = new OfficeAgentSession();
  let count = 0;
  const fetchImpl = vi.fn(async () => ++count <= 3 ? tool(String(count), "list_tables", {}) : done());
  await session.run("key", "Explore", ctx, undefined, { fetchImpl });
  expect(ctx.listTables).toHaveBeenCalledTimes(3); // initial inventory plus two tool calls
  expect(JSON.stringify(session.snapshot())).toContain("already called");
  count = 0; await session.run("key", "Explore again", ctx, undefined, { fetchImpl });
  expect(ctx.listTables).toHaveBeenCalledTimes(5);
});
it("stops executing SQL after three different failures and preserves earlier results", async () => {
  const ctx = context(), session = new OfficeAgentSession(), onQueryResult = vi.fn();
  ctx.backend.query = vi.fn().mockResolvedValueOnce(empty).mockRejectedValue(new Error("Missing column"));
  let count = 0;
  const fetchImpl = vi.fn(async () => ++count <= 5 ? tool(String(count), "run_sql", { sql: `SELECT ${count}` }) : done());
  await session.run("key", "Calculate", ctx, undefined, { fetchImpl }, { onQueryResult });
  expect(ctx.backend.query).toHaveBeenCalledTimes(4);
  expect(onQueryResult).toHaveBeenCalledTimes(1);
  expect(JSON.stringify(session.snapshot())).toContain("three-query failure budget");
});
