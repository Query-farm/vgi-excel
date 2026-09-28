import { CLARIFICATION_TOOL, RESULT_SCOPE_SCHEMA } from "./agent-experience.js";
import { CATALOG_AGENT_TOOLS, CATALOG_TOOL_NAMES } from "./agent-catalog.js";
import { assertAgentReadOnlySql } from "./sql.js";
import type { QueryBackend, QueryResult } from "./types.js";

export interface AgentToolContext {
  backend: QueryBackend;
  catalogTool?(name: string, input: Record<string, unknown>, signal?: AbortSignal): Promise<string>;
  listTables(signal?: AbortSignal): Promise<unknown>;
  listFunctions(signal?: AbortSignal): Promise<unknown>;
  describeTable(input: { catalog?: string; schema: string; table: string }, signal?: AbortSignal): Promise<unknown>;
  createQuery?(input: { name: string; sql: string }): Promise<void> | void;
}

export interface AgentToolResult {
  content: string;
  queryResult?: QueryResult;
}

export async function executeAgentTool(
  name: string,
  input: Record<string, unknown>,
  context: AgentToolContext,
  signal?: AbortSignal,
): Promise<AgentToolResult> {
  signal?.throwIfAborted();
  if (context.catalogTool && CATALOG_TOOL_NAMES.has(name)) return { content: await context.catalogTool(name, input, signal) };
  switch (name) {
    case "run_sql": {
      const sql = String(input.sql ?? "");
      assertAgentReadOnlySql(sql);
      const queryResult = await context.backend.query(sql, { maxRows: 10_000, signal });
      return { content: summarizeResult(queryResult), queryResult };
    }
    case "list_tables":
      return { content: JSON.stringify(await context.listTables(signal)) };
    case "list_functions":
      return { content: JSON.stringify(await context.listFunctions(signal)) };
    case "describe_table":
      return {
        content: JSON.stringify(
          await context.describeTable({
            catalog: input.catalog ? String(input.catalog) : undefined,
            schema: String(input.schema ?? ""),
            table: String(input.table ?? ""),
          }, signal),
        ),
      };
    case "create_query_tab": {
      const name = String(input.name ?? "").trim();
      const sql = String(input.sql ?? "").trim();
      if (!name) throw new Error("A descriptive query tab name is required.");
      if (name.length > 120) throw new Error("The query tab name must be 120 characters or fewer.");
      assertAgentReadOnlySql(sql);
      if (!context.createQuery) throw new Error("Creating Query Editor tabs is unavailable in this host.");
      await context.createQuery({ name, sql });
      return { content: JSON.stringify({ created: true, name, message: "The SQL was saved in a new Query Editor tab. It has not been executed." }) };
    }
    default:
      throw new Error(`Unknown agent tool: ${name}`);
  }
}

export function summarizeResult(result: QueryResult): string {
  return JSON.stringify({
    columns: result.columns,
    rows: result.rows.slice(0, 20),
    row_count: result.rowCount,
    sample_row_count: Math.min(20, result.rows.length),
    truncated: result.rows.length > 20 || !!result.truncated || result.rows.length < result.rowCount,
  });
}

const BASE_AGENT_TOOLS = [
  {
    name: "run_sql",
    description: "Run one read-only SQL statement. Returns columns, up to 20 sample rows, and the total row count.",
    input_schema: { type: "object", properties: { sql: { type: "string" }, scope: RESULT_SCOPE_SCHEMA }, required: ["sql"] },
  },
  {
    name: "list_tables",
    description: "List available catalogs, schemas, tables, and views.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "list_functions",
    description: "List every callable with qualified names plus rich VGI named/positional arguments, constraints, and documentation.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "describe_table",
    description: "Describe a table or view and its columns.",
    input_schema: {
      type: "object",
      properties: {
        catalog: { type: "string" },
        schema: { type: "string" },
        table: { type: "string" },
      },
      required: ["schema", "table"],
    },
  },
  {
    name: "create_query_tab",
    description: "Create a new locally saved Query Editor tab containing a useful read-only SQL query. Use this when the user asks to save, keep, edit, or open a query. This does not execute the SQL or modify the workbook.",
    input_schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        name: { type: "string", description: "A concise descriptive name for the query tab." },
        sql: { type: "string", description: "One complete read-only SQL query using the active attached catalog." },
      },
      required: ["name", "sql"],
    },
  },
] as const;

export const AGENT_TOOLS = [CLARIFICATION_TOOL, ...BASE_AGENT_TOOLS.filter(tool => !CATALOG_TOOL_NAMES.has(tool.name)), ...CATALOG_AGENT_TOOLS];
