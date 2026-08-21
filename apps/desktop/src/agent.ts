import { agentRoundSeparator, appendAgentRoundText, assertAgentReadOnlySql, buildExcelAgentSystemPrompt, excelWorksheetName, type AgentCatalogObject, type QueryResult } from "@query-farm/vgi-excel-core";
import { host, type DesktopConnection } from "./bridge";
import { sanitizeConversation } from "./agent-history";
import { recordToolCall, repeatedCallMessage } from "./agent-loop-guard";
import { parseStreamedToolInput } from "./tool-input";

export type AgentMessage = { role: "user" | "assistant"; content: string | AgentBlock[] };
export type AgentBlock = TextBlock | ToolUseBlock | ToolResultBlock;
export type TextBlock = { type: "text"; text: string };
export type ToolUseBlock = { type: "tool_use"; id: string; name: string; input: Record<string, unknown> };
export type ToolResultBlock = { type: "tool_result"; tool_use_id: string; content: string; is_error?: boolean };

export interface AgentCallbacks {
  onText(chunk: string): void;
  onTool(name: string, state: "writing" | "running" | "done" | "error", detail?: string, callId?: string): void;
  onRetry(message: string | null): void;
  onResult(result: QueryResult): void;
  onWorkbookAction(action: WorkbookWriteAction): void;
  onQueryDocument(value: { name: string; sql: string }): void;
}

export interface WorkbookWriteAction {
  id: string;
  mode: "new_sheet" | "replace_table";
  result: QueryResult;
  sheetName?: string;
  tableName: string;
}

const TOOLS = [
  { name: "run_sql", description: "Execute one read-only SQL query. Returns columns, the first 20 rows, total row count, and a result ID for paging.", input_schema: { type: "object", additionalProperties: false, properties: { sql: { type: "string" } }, required: ["sql"] } },
  { name: "read_query_results", description: "Read more rows from an earlier run_sql result without executing it again.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, offset: { type: "number" }, limit: { type: "number" } }, required: ["result_id"] } },
  { name: "list_tables", description: "List catalogs, schemas, tables, and views available through the selected VGI connection.", input_schema: { type: "object", additionalProperties: false, properties: {} } },
  { name: "list_functions", description: "Inspect scalar, macro, and table-function signatures before calling them. Filter by catalog, schema, or function name when possible.", input_schema: { type: "object", additionalProperties: false, properties: { catalog: { type: "string" }, schema: { type: "string" }, name: { type: "string", description: "Exact or partial function name." } } } },
  { name: "describe_table", description: "Describe the columns of a table or view before querying it.", input_schema: { type: "object", additionalProperties: false, properties: { catalog: { type: "string" }, schema: { type: "string" }, table: { type: "string" } }, required: ["schema", "table"] } },
  { name: "create_query_tab", description: "Create a new locally saved Query Editor tab containing one useful read-only SQL query. Use this when the user asks to save, keep, edit, or open a query. This does not execute SQL or modify the workbook.", input_schema: { type: "object", additionalProperties: false, properties: { name: { type: "string", description: "A concise descriptive tab name." }, sql: { type: "string", description: "One complete read-only SQL query using the active attached catalog." } }, required: ["name", "sql"] } },
  { name: "workbook_overview", description: "Inspect the active workbook, worksheets, Excel tables, current selection, and formula counts. Read-only.", input_schema: { type: "object", additionalProperties: false, properties: {} } },
  { name: "read_range", description: "Read values and formulas from a bounded A1 range in the active workbook. Read-only; use a range no larger than 5,000 cells.", input_schema: { type: "object", additionalProperties: false, properties: { sheet: { type: "string" }, address: { type: "string", description: "A1 range such as A1:F40, without a sheet prefix." } }, required: ["sheet", "address"] } },
  { name: "list_formulas", description: "List formulas and their current displayed values on one worksheet or across the active workbook. Read-only.", input_schema: { type: "object", additionalProperties: false, properties: { sheet: { type: "string" }, limit: { type: "number" } } } },
  { name: "stage_result_to_new_sheet", description: "Stage a prior run_sql result for insertion as a formatted Excel table on a new worksheet. This does not modify Excel until the user confirms the action.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, sheet_name: { type: "string", maxLength: 31, description: "A concise Excel worksheet name. Do not use \\ / ? * [ ] : characters." }, table_name: { type: "string" } }, required: ["result_id", "sheet_name", "table_name"] } },
  { name: "stage_result_to_table", description: "Stage a prior run_sql result to replace the data in an existing Excel table. This does not modify Excel until the user confirms the action.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, table_name: { type: "string" } }, required: ["result_id", "table_name"] } },
] as const;

const LIST_TABLES = `SELECT table_catalog, table_schema, table_name, table_type FROM information_schema.tables WHERE table_schema NOT IN ('information_schema','pg_catalog') ORDER BY 1,2,3`;
const LIST_FUNCTIONS = `SELECT database_name, schema_name, function_name, function_type, CAST(to_json(parameters) AS VARCHAR) AS parameters, CAST(to_json(parameter_types) AS VARCHAR) AS parameter_types, return_type, description, comment, CAST(to_json(examples) AS VARCHAR) AS examples, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_functions() WHERE database_name NOT IN ('system','temp')`;
const LIST_VGI_ARGUMENTS = `SELECT catalog_name, schema_name, function_name, function_type, arg_position, field_index, arg_name, arg_type, arg_description, is_named, is_positional, is_const, is_varargs, is_table_input, is_any_type, arg_default, arg_choices, arg_range, arg_pattern FROM vgi_function_arguments() WHERE true`;
const CATALOG_CONTEXT = (catalog: string) => `SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, CASE WHEN table_type='VIEW' THEN 'view' ELSE 'table' END AS kind, '' AS description FROM information_schema.tables WHERE table_catalog=${literal(catalog)} AND table_schema NOT IN ('information_schema','pg_catalog') UNION ALL SELECT database_name, schema_name, function_name, CASE WHEN function_type IN ('macro','table_macro') THEN 'macro' ELSE function_type END, COALESCE(description, comment, '') FROM duckdb_functions() WHERE database_name=${literal(catalog)} ORDER BY 1,2,4,3`;

export const DEFAULT_MODEL = "claude-sonnet-4-6";
export class AgentSession {
  readonly messages: AgentMessage[] = [];
  private readonly results = new Map<string, QueryResult>();
  private readonly prompts = new Map<string, string>();

  reset(): void { this.messages.length = 0; this.results.clear(); this.prompts.clear(); }
  restore(messages: AgentMessage[]): void {
    this.reset();
    this.messages.push(...JSON.parse(JSON.stringify(messages)) as AgentMessage[]);
    sanitizeConversation(this.messages);
  }
  snapshot(): AgentMessage[] { return JSON.parse(JSON.stringify(this.messages)) as AgentMessage[]; }

  async run(apiKey: string, model: string, prompt: string, connection: DesktopConnection, callbacks: AgentCallbacks, signal: AbortSignal): Promise<void> {
    if (!apiKey.trim()) throw new Error("An Anthropic API key is required.");
    if (!prompt.trim()) throw new Error("Enter a question for the agent.");
    this.messages.push({ role: "user", content: prompt.trim() });
    const repeated = new Map<string, number>();
    const runId = crypto.randomUUID();
    let sqlFailures = 0;
    let visibleText = "";
    trace({ event: "run_start", run_id: runId, model: model || DEFAULT_MODEL, connection: connection.name, catalog: connection.catalog, prompt_chars: prompt.trim().length });

    try {
      const systemPrompt = await this.systemPrompt(connection);
      for (let round = 0; round < 20; round++) {
        if (signal.aborted) throw new DOMException("Stopped", "AbortError");
        sanitizeConversation(this.messages);
        trace({ event: "round_start", run_id: runId, round: round + 1, message_count: this.messages.length });
        let firstTextChunk = true;
        const roundCallbacks: AgentCallbacks = { ...callbacks, onText: (chunk) => {
          if (firstTextChunk && chunk) { firstTextChunk = false; const separator = agentRoundSeparator(visibleText, chunk); if (separator) callbacks.onText(separator); }
          callbacks.onText(chunk);
        } };
        const blocks = await streamRequest(apiKey, model || DEFAULT_MODEL, systemPrompt, this.messages, roundCallbacks, signal);
        const roundText = blocks.filter((block): block is TextBlock => block.type === "text").map((block) => block.text).join("");
        visibleText = appendAgentRoundText(visibleText, roundText);
        this.messages.push({ role: "assistant", content: blocks });
        const calls = blocks.filter((block): block is ToolUseBlock => block.type === "tool_use");
        if (!calls.length) {
          trace({ event: "run_complete", run_id: runId, rounds: round + 1, response_chars: blocks.filter((item): item is TextBlock => item.type === "text").reduce((sum, item) => sum + item.text.length, 0) });
          return;
        }
        try {
          const results: ToolResultBlock[] = [];
          for (const call of calls) {
            if (signal.aborted) throw new DOMException("Stopped", "AbortError");
            callbacks.onTool(call.name, "running", summarizeInput(call.input), call.id);
            trace({ event: "tool_call", run_id: runId, round: round + 1, tool_use_id: call.id, tool: call.name, input: call.input });
            const parsedError = typeof call.input.__parseError === "string";
            const repeat = recordToolCall(repeated, call.name, call.input);
            let blocked: string | null = null;
            if (parsedError) blocked = `The streamed ${call.name} arguments were not valid JSON. Recreate the tool call with a valid JSON object.`;
            else if (repeat.block) blocked = repeatedCallMessage(call.name, repeat.count);
            else if (call.name === "run_sql" && sqlFailures >= 3) blocked = "The three-query failure budget for this turn has been reached. Explain the missing function signature or metadata instead of running another query.";
            if (blocked) {
              callbacks.onTool(call.name, "error", blocked, call.id);
              trace({ event: "tool_blocked", run_id: runId, round: round + 1, tool_use_id: call.id, tool: call.name, reason: blocked });
              results.push({ type: "tool_result", tool_use_id: call.id, content: blocked, is_error: true });
              continue;
            }
            const started = performance.now();
            try {
              const content = await this.execute(call.name, call.input, connection.name, callbacks);
              callbacks.onTool(call.name, "done", completedToolDetail(call.name, call.input, content), call.id);
              trace({ event: "tool_result", run_id: runId, round: round + 1, tool_use_id: call.id, tool: call.name, elapsed_ms: Math.round(performance.now() - started), result_chars: content.length });
              results.push({ type: "tool_result", tool_use_id: call.id, content });
            } catch (error) {
              if ((error as Error).name === "AbortError") throw error;
              let content = error instanceof Error ? error.message : String(error);
              if (call.name === "run_sql") {
                sqlFailures++;
                content += `\n\nQuery failure ${sqlFailures} of 3 for this turn.${sqlFailures >= 3 ? " Do not run another query; explain what signature or metadata is missing." : " Inspect function metadata and change the approach before retrying."}`;
              }
              callbacks.onTool(call.name, "error", failedToolDetail(call.name, call.input, content), call.id);
              trace({ event: "tool_error", run_id: runId, round: round + 1, tool_use_id: call.id, tool: call.name, elapsed_ms: Math.round(performance.now() - started), error: content });
              results.push({ type: "tool_result", tool_use_id: call.id, content, is_error: true });
            }
          }
          this.messages.push({ role: "user", content: results });
        } catch (error) {
          if ((error as Error).name === "AbortError") {
            const text = blocks.filter((block): block is TextBlock => block.type === "text" && !!block.text);
            this.messages[this.messages.length - 1] = { role: "assistant", content: text.length ? text : [{ type: "text", text: "(stopped)" }] };
          }
          throw error;
        }
      }
      throw new Error("The agent reached the 20-round safety limit.");
    } catch (error) {
      trace({ event: (error as Error).name === "AbortError" ? "run_stopped" : "run_error", run_id: runId, error: error instanceof Error ? error.message : String(error) });
      throw error;
    }
  }

  private async systemPrompt(connection: DesktopConnection): Promise<string> {
    const cacheKey = `${connection.name}\0${connection.catalog}\0${connection.authentication}`;
    const cached = this.prompts.get(cacheKey);
    if (cached) return cached;
    let objects: AgentCatalogObject[] = [];
    let inventoryError: string | undefined;
    try {
      const result = await host.query(CATALOG_CONTEXT(connection.catalog), connection.name, true, 20_000);
      objects = records(result).map((row) => ({
        catalog: String(row.catalog ?? connection.catalog),
        schema: String(row.schema ?? "main"),
        name: String(row.name ?? ""),
        kind: String(row.kind ?? "object"),
        description: row.description == null ? undefined : String(row.description),
      })).filter((item) => item.name);
    } catch (error) {
      inventoryError = error instanceof Error ? error.message : String(error);
    }
    const value = buildExcelAgentSystemPrompt({
      connection: { name: connection.name, catalog: connection.catalog, authentication: connection.authentication },
      objects,
      inventoryError,
    });
    this.prompts.set(cacheKey, value);
    return value;
  }

  private async execute(name: string, input: Record<string, unknown>, connection: string, callbacks: AgentCallbacks): Promise<string> {
    if (name === "run_sql") {
      const result = await host.query(String(input.sql ?? ""), connection, true);
      const id = crypto.randomUUID();
      this.results.set(id, result);
      callbacks.onResult(result);
      return JSON.stringify({ result_id: id, columns: result.columns, rows: result.rows.slice(0, 20), row_count: result.rowCount, truncated: result.rows.length > 20 || result.truncated });
    }
    if (name === "read_query_results") {
      const result = this.results.get(String(input.result_id ?? ""));
      if (!result) throw new Error("That query result is no longer available.");
      const offset = Math.max(0, Number(input.offset ?? 0));
      const limit = Math.max(1, Math.min(100, Number(input.limit ?? 20)));
      return JSON.stringify({ columns: result.columns, rows: result.rows.slice(offset, offset + limit), offset, row_count: result.rowCount });
    }
    if (name === "list_tables") return summarize(await host.query(LIST_TABLES, connection, true, 10_000));
    if (name === "list_functions") {
      const flatFilters: string[] = [], richFilters: string[] = [];
      if (input.catalog) { flatFilters.push(`database_name=${literal(String(input.catalog))}`); richFilters.push(`catalog_name=${literal(String(input.catalog))}`); }
      if (input.schema) { flatFilters.push(`schema_name=${literal(String(input.schema))}`); richFilters.push(`schema_name=${literal(String(input.schema))}`); }
      if (input.name) { const match = literal(`%${String(input.name)}%`); flatFilters.push(`function_name ILIKE ${match}`); richFilters.push(`function_name ILIKE ${match}`); }
      const flatSql = `${LIST_FUNCTIONS}${flatFilters.length ? ` AND ${flatFilters.join(" AND ")}` : ""} ORDER BY 1,2,3`;
      const richSql = `${LIST_VGI_ARGUMENTS}${richFilters.length ? ` AND ${richFilters.join(" AND ")}` : ""} ORDER BY 1,2,3,field_index`;
      const [functions, argumentsResult] = await Promise.all([
        host.query(flatSql, connection, true, 10_000),
        host.query(richSql, connection, true, 10_000),
      ]);
      return JSON.stringify({
        guidance: "Use positional arguments in arg_position order. Pass every argument with is_named=true as name := value. Respect choices, ranges, patterns, and exact types; cast numeric literals when needed.",
        functions: combineFunctionMetadata(functions, argumentsResult),
      });
    }
    if (name === "describe_table") {
      const filters = [`table_schema=${literal(String(input.schema ?? ""))}`, `table_name=${literal(String(input.table ?? ""))}`];
      if (input.catalog) filters.push(`table_catalog=${literal(String(input.catalog))}`);
      return summarize(await host.query(`SELECT table_catalog, table_schema, table_name, column_name, data_type, is_nullable FROM information_schema.columns WHERE ${filters.join(" AND ")} ORDER BY ordinal_position`, connection, true));
    }
    if (name === "create_query_tab") {
      const queryName = String(input.name ?? "").trim(), sql = String(input.sql ?? "").trim();
      if (!queryName) throw new Error("A descriptive query tab name is required.");
      if (queryName.length > 120) throw new Error("The query tab name must be 120 characters or fewer.");
      assertAgentReadOnlySql(sql);
      callbacks.onQueryDocument({ name: queryName, sql });
      return JSON.stringify({ created: true, name: queryName, message: "The SQL was saved in a new Query Editor tab. It has not been executed." });
    }
    if (name === "workbook_overview") return JSON.stringify(await host.workbookOverview());
    if (name === "read_range") return JSON.stringify(await host.readRange(String(input.sheet ?? ""), String(input.address ?? "")));
    if (name === "list_formulas") return JSON.stringify(await host.listFormulas(input.sheet ? String(input.sheet) : undefined, Number(input.limit ?? 200)));
    if (name === "stage_result_to_new_sheet" || name === "stage_result_to_table") {
      const resultId = String(input.result_id ?? "");
      const result = this.results.get(resultId);
      if (!result) throw new Error("That query result is no longer available. Run the query again before staging a workbook action.");
      const action: WorkbookWriteAction = {
        id: crypto.randomUUID(),
        mode: name === "stage_result_to_new_sheet" ? "new_sheet" : "replace_table",
        result,
        sheetName: name === "stage_result_to_new_sheet" ? excelWorksheetName(input.sheet_name) : undefined,
        tableName: String(input.table_name ?? "VGI_Result"),
      };
      callbacks.onWorkbookAction(action);
      return JSON.stringify({ staged: true, action_id: action.id, mode: action.mode, sheet_name: action.sheetName, table_name: action.tableName, row_count: result.rowCount, message: "The workbook has not changed. The user must confirm this action in the Workbench." });
    }
    throw new Error(`Unknown agent tool: ${name}`);
  }
}

async function streamRequest(apiKey: string, model: string, systemPrompt: string, messages: AgentMessage[], callbacks: AgentCallbacks, signal: AbortSignal): Promise<Array<TextBlock | ToolUseBlock>> {
  const response = await fetchWithRetry("https://api.anthropic.com/v1/messages", {
    method: "POST", signal,
    headers: { "content-type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
    body: JSON.stringify({ model, max_tokens: 8192, stream: true, system: systemPrompt, tools: TOOLS, messages }),
  }, callbacks);
  if (!response.body) throw new Error("Anthropic returned an empty response stream.");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "", current: TextBlock | ToolUseBlock | null = null, toolJson = "";
  const blocks: Array<TextBlock | ToolUseBlock> = [];
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n"); buffer = lines.pop() ?? "";
    for (const line of lines) {
      if (!line.startsWith("data: ")) continue;
      const raw = line.slice(6).trim(); if (!raw || raw === "[DONE]") continue;
      let event: any; try { event = JSON.parse(raw); } catch { continue; }
      if (event.type === "content_block_start") {
        if (event.content_block?.type === "text") current = { type: "text", text: "" };
        if (event.content_block?.type === "tool_use") {
          current = { type: "tool_use", id: event.content_block.id, name: event.content_block.name, input: {} };
          toolJson = ""; callbacks.onTool(current.name, "writing", undefined, current.id);
        }
      } else if (event.type === "content_block_delta" && current) {
        if (event.delta?.type === "text_delta" && current.type === "text") { current.text += event.delta.text; callbacks.onText(event.delta.text); }
        if (event.delta?.type === "input_json_delta" && current.type === "tool_use") toolJson += event.delta.partial_json;
      } else if (event.type === "content_block_stop" && current) {
        if (current.type === "tool_use") {
          current.input = parseStreamedToolInput(toolJson).input;
        }
        blocks.push(current); current = null;
      } else if (event.type === "error") {
        throw new Error(event.error?.message || "Anthropic streaming error.");
      }
    }
  }
  return blocks;
}

async function fetchWithRetry(url: string, init: RequestInit, callbacks: AgentCallbacks, retries = 3): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    try {
      const response = await fetch(url, init);
      if (response.ok) { callbacks.onRetry(null); return response; }
      const body = await response.text();
      if ((response.status !== 429 && response.status !== 529) || attempt >= retries) throw new Error(`Anthropic request failed (${response.status}): ${body}`);
      const seconds = Math.min(30, Math.max(1, Number(response.headers.get("retry-after")) || 2 ** attempt));
      callbacks.onRetry(`Anthropic is busy; retrying in ${seconds}s…`);
      await delay(seconds * 1000, init.signal);
    } catch (error) {
      if (init.signal?.aborted) throw new DOMException("Stopped", "AbortError");
      if (!(error instanceof TypeError) || attempt >= retries) throw error;
      const seconds = Math.min(10, 2 ** attempt);
      callbacks.onRetry(`Network interruption; retrying in ${seconds}s…`);
      await delay(seconds * 1000, init.signal);
    }
  }
}

function delay(ms: number, signal?: AbortSignal | null): Promise<void> { return new Promise((resolve, reject) => { const timer = setTimeout(resolve, ms); signal?.addEventListener("abort", () => { clearTimeout(timer); reject(new DOMException("Stopped", "AbortError")); }, { once: true }); }); }
function summaryObject(result: QueryResult): Record<string, unknown> { return { columns: result.columns, rows: result.rows.slice(0, 100), row_count: result.rowCount, truncated: result.truncated || result.rows.length > 100 }; }
function summarize(result: QueryResult): string { return JSON.stringify(summaryObject(result)); }
function combineFunctionMetadata(functions: QueryResult, argumentsResult: QueryResult): Record<string, unknown>[] {
  const argumentsByFunction = new Map<string, Record<string, unknown>[]>();
  for (const row of records(argumentsResult)) {
    const key = functionKey(row.catalog_name, row.schema_name, row.function_name);
    const item = {
      name: row.arg_name,
      type: row.arg_type,
      position: row.arg_position,
      kind: truthy(row.is_named) ? "named" : truthy(row.is_positional) ? "positional" : truthy(row.is_varargs) ? "varargs" : "other",
      description: row.arg_description,
      default: parseJson(row.arg_default),
      choices: parseJson(row.arg_choices),
      range: row.arg_range,
      pattern: row.arg_pattern,
      const: truthy(row.is_const),
      table_input: truthy(row.is_table_input),
      any_type: truthy(row.is_any_type),
    };
    const values = argumentsByFunction.get(key) ?? [];
    values.push(Object.fromEntries(Object.entries(item).filter(([, value]) => value !== null && value !== undefined && value !== false)));
    argumentsByFunction.set(key, values);
  }
  return records(functions).map((row) => {
    const rich = argumentsByFunction.get(functionKey(row.database_name, row.schema_name, row.function_name));
    const names = stringList(row.parameters), types = stringList(row.parameter_types);
    const fallback = types.map((type, index) => ({ name: names[index] ?? `arg_${index + 1}`, type, kind: "unknown" }));
    return Object.fromEntries(Object.entries({
      catalog: row.database_name,
      schema: row.schema_name,
      name: row.function_name,
      type: row.function_type,
      description: row.description ?? row.comment,
      return_type: row.return_type,
      arguments: rich ?? fallback,
      examples: parseJson(row.examples),
      tags: parseJson(row.tags),
    }).filter(([, value]) => value !== null && value !== undefined && value !== ""));
  });
}
function records(result: QueryResult): Record<string, unknown>[] { return result.rows.map((row) => Object.fromEntries(result.columns.map((column, index) => [column.name, row[index]]))); }
function functionKey(catalog: unknown, schema: unknown, name: unknown): string { return `${String(catalog ?? "")}\0${String(schema ?? "")}\0${String(name ?? "")}`; }
function stringList(value: unknown): string[] { const parsed = parseJson(value); return Array.isArray(parsed) ? parsed.map(String) : []; }
function parseJson(value: unknown): unknown { if (typeof value !== "string") return value ?? null; if (!value.trim()) return null; try { return JSON.parse(value); } catch { return value; } }
function truthy(value: unknown): boolean { return value === true || value === 1 || String(value).toLowerCase() === "true"; }
function literal(value: string): string { return `'${value.replaceAll("'", "''")}'`; }
function summarizeInput(input: Record<string, unknown>): string { return typeof input.sql === "string" ? input.sql : JSON.stringify(input); }
function preview(value: string): string { return value.length > 240 ? `${value.slice(0, 240)}…` : value; }
function completedToolDetail(name: string, input: Record<string, unknown>, content: string): string {
  return name === "run_sql" && typeof input.sql === "string" ? input.sql : preview(content);
}
function failedToolDetail(name: string, input: Record<string, unknown>, error: string): string {
  return name === "run_sql" && typeof input.sql === "string" ? `${input.sql}\n\nError: ${error}` : error;
}
function trace(event: Record<string, unknown>): void {
  const safe = redactTrace(event) as Record<string, unknown>;
  const label = `[VGI agent] ${String(safe.event ?? "event")}`;
  if (typeof console.groupCollapsed === "function") { console.groupCollapsed(label); console.info(safe); console.groupEnd(); }
  else console.info(label, safe);
  void host.traceAgent(safe).catch(() => undefined);
}
function redactTrace(value: unknown, key = ""): unknown {
  if (/api.?key|authorization|credential|secret|token/i.test(key)) return "***";
  if (Array.isArray(value)) return value.map((item) => redactTrace(item));
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([name, item]) => [name, redactTrace(item, name)]));
  if (typeof value !== "string") return value;
  return value
    .replace(/sk-ant-[A-Za-z0-9_-]+/gi, "sk-ant-***")
    .replace(/((?:api[_-]?key|bearer_token|authorization|secret)\s*(?::=|=>|=|:)\s*['"]?)[^'"\s,;]+/gi, "$1***");
}
