import { describe, expect, it } from "vitest";
import { executeCatalogTool } from "./agent-catalog.js";
import { assertAgentReadOnlySql } from "./sql.js";
import type { QueryResult } from "./types.js";
function result(rows: Record<string, unknown>[]): QueryResult {
  const columns = Object.keys(rows[0] ?? {}).map(name => ({ name, type: "VARCHAR" }));
  return { columns, rows: rows.map(row => columns.map(col => row[col.name] as string)), rowCount: rows.length };
}

describe("VGI AI discovery ports", () => {
  it("explains required filters on views, decodes JSON tags, and deduplicates examples", async () => {
    const queries: string[] = [];
    const answer = JSON.parse(await executeCatalogTool("describe_table", { catalog: "accounts", schema: "main", table: "sales'2026" }, async sql => {
      assertAgentReadOnlySql(sql); queries.push(sql);
      if (sql.includes("duckdb_columns")) return result([{ catalog: "accounts", schema: "main", name: "sales'2026", column_name: "country", data_type: "VARCHAR" }]);
      if (sql.includes("duckdb_constraints")) return result([{ catalog: "accounts", schema: "main", name: "sales'2026", constraint_type: "PRIMARY KEY", columns: '["id"]' }]);
      return result([{ catalog: "accounts", schema: "main", name: "sales'2026", kind: "view", tags: JSON.stringify({ vgi_required_filters: '[["country"],["start","end"]]', "vgi.keywords": '["sales","revenue"]', "vgi.example_queries": '[{"sql":"SELECT 1"}]', "vgi.executable_examples": '[{"sql":"SELECT   1"}]', "vgi.agent_test_tasks": "private-grader" }) }]);
    }));
    expect(answer.objects[0]).toMatchObject({ required_filters: [["country"], ["start", "end"]], required_filters_rule: expect.stringContaining("AND of OR-groups"), tags: { "vgi.keywords": ["sales", "revenue"] }, examples: [{ description: null, sql: "SELECT 1" }] });
    expect(answer.objects[0].columns).toHaveLength(1);
    expect(answer.objects[0].constraints).toHaveLength(1);
    expect(JSON.stringify(answer)).not.toContain("private-grader");
    expect(queries.every(sql => sql.includes("'sales''2026'"))).toBe(true);
  });
  it("keeps zero-argument functions and labels unavailable rich metadata", async () => {
    const answer = JSON.parse(await executeCatalogTool("describe_function", { schema: "main", function: "now" }, async sql => {
      assertAgentReadOnlySql(sql);
      if (sql.includes("vgi_function_arguments")) throw new Error("Unavailable");
      return result([{ catalog: "weather", schema: "main", name: "now", kind: "scalar", parameters: "[]", parameter_types: "[]" }]);
    }));
    expect(answer.objects[0]).toMatchObject({ qualified_name: "weather.main.now", arguments: [] });
    expect(answer.warning).toContain("unavailable");
  });
  it("paginates filtered objects and bounds corrupt limits", async () => {
    const query = async () => result(["alpha", "beta", "gamma"].map(name => ({ catalog: "c", schema: "main", name, tags: '{"vgi.category":"finance"}' })));
    const first = JSON.parse(await executeCatalogTool("list_tables", { category: "finance", limit: 1 }, query));
    expect(first).toMatchObject({ total: 3, next_cursor: "1" });
    const second = JSON.parse(await executeCatalogTool("list_tables", { cursor: first.next_cursor, limit: 1 }, query));
    expect(second.objects[0].name).toBe("beta");
    expect(JSON.parse(await executeCatalogTool("list_tables", { limit: "bad" }, query)).objects).toHaveLength(3);
  });
});
