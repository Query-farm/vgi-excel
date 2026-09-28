import { recordAITool } from "@query-farm/vgi-excel-core";
import { CLARIFICATION_TOOL, RESULT_SCOPE_SCHEMA, clarificationInput, resultScope, type AgentClarification, type AgentResultCard } from "@query-farm/vgi-excel-core";
import { agentRoundSeparator, appendAgentRoundText, assertAgentReadOnlySql, buildExcelAgentSystemPrompt, excelWorksheetName, type AgentCatalogObject, type QueryResult } from "@query-farm/vgi-excel-core";
import { host, type DesktopConnection } from "./bridge";
import { sanitizeConversation } from "./agent-history";
import { recordToolCall, repeatedCallMessage } from "./agent-loop-guard";

export type { AgentMessage, AgentBlock, TextBlock, ToolUseBlock, ToolResultBlock } from "@query-farm/vgi-excel-core";
import { CATALOG_AGENT_TOOLS, CATALOG_TOOL_NAMES, executeCatalogTool, DEFAULT_AI_MODEL, requestAnthropic, repairAgentHistory, type AIRequestOptions, type AgentMessage, type AgentBlock, type TextBlock, type ToolUseBlock, type ToolResultBlock } from "@query-farm/vgi-excel-core";

export interface AgentCallbacks {
  onThinking?(text: string): void;
  onStage?(stage: string): void;
  onClarification?(question: AgentClarification, signal?: AbortSignal): Promise<string>;
  onQueryResult?(value: AgentResultCard): void;
  onText(chunk: string): void;
  onTool(name: string, state: "writing" | "running" | "done" | "error", detail?: string, callId?: string): void;
  onRetry(message: string | null): void;
  onResult(result: QueryResult): void;
  onWorkbookAction(action: WorkbookWriteAction): void;
  onQueryDocument(value: { name: string; sql: string }): void;
}

export interface WorkbookWriteAction {
  id: string;
  resultId?: string;
  mode: "new_sheet" | "replace_table";
  result: QueryResult;
  query?: { sql: string; connection: string };
  sheetName?: string;
  tableName: string;
}

const BASE_TOOLS = [
  { name: "run_sql", description: "Execute one read-only SQL query. Returns columns, the first 20 rows, total row count, and a result ID for paging.", input_schema: { type: "object", additionalProperties: false, properties: { sql: { type: "string" }, scope: RESULT_SCOPE_SCHEMA }, required: ["sql"] } },
  { name: "read_query_results", description: "Read more rows from an earlier run_sql result without executing it again.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, offset: { type: "number" }, limit: { type: "number" } }, required: ["result_id"] } },
  { name: "list_tables", description: "List catalogs, schemas, tables, and views available through the selected VGI connection.", input_schema: { type: "object", additionalProperties: false, properties: {} } },
  { name: "list_functions", description: "Inspect scalar, macro, and table-function signatures before calling them. Filter by catalog, schema, or function name when possible.", input_schema: { type: "object", additionalProperties: false, properties: { catalog: { type: "string" }, schema: { type: "string" }, name: { type: "string", description: "Exact or partial function name." } } } },
  { name: "describe_table", description: "Describe the columns of a table or view before querying it.", input_schema: { type: "object", additionalProperties: false, properties: { catalog: { type: "string" }, schema: { type: "string" }, table: { type: "string" } }, required: ["schema", "table"] } },
  { name: "create_query_tab", description: "Create a new locally saved Query Editor tab containing one useful read-only SQL query. Use this when the user asks to save, keep, edit, or open a query. This does not execute SQL or modify the workbook.", input_schema: { type: "object", additionalProperties: false, properties: { name: { type: "string", description: "A concise descriptive tab name." }, sql: { type: "string", description: "One complete read-only SQL query using the active attached catalog." } }, required: ["name", "sql"] } },
  { name: "workbook_overview", description: "Inspect the active workbook, worksheets, Excel tables, current selection, and formula counts. Read-only.", input_schema: { type: "object", additionalProperties: false, properties: {} } },
  { name: "read_range", description: "Read values and formulas from a bounded A1 range in the active workbook. Read-only; use a range no larger than 5,000 cells.", input_schema: { type: "object", additionalProperties: false, properties: { sheet: { type: "string" }, address: { type: "string", description: "A1 range such as A1:F40, without a sheet prefix." } }, required: ["sheet", "address"] } },
  { name: "list_formulas", description: "List formulas and their current displayed values on one worksheet or across the active workbook. Read-only.", input_schema: { type: "object", additionalProperties: false, properties: { sheet: { type: "string" }, limit: { type: "number" } } } },
  { name: "stage_result_to_new_sheet", description: "Stage the original SQL of a prior run_sql result as a refreshable Excel Power Query table on a new worksheet. Use this by default when placing data in Excel. Loading reruns the full query through the saved connection; preview rows are not copied. This does not modify Excel until the user confirms the action.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, sheet_name: { type: "string", maxLength: 31, description: "A concise Excel worksheet name. Do not use \\ / ? * [ ] : characters." }, table_name: { type: "string" } }, required: ["result_id", "sheet_name", "table_name"] } },
  { name: "stage_result_to_table", description: "Stage a static snapshot replacement only when the user explicitly requests static data in an existing table. This does not create a refresh connection. For normal Excel output, use stage_result_to_new_sheet to create a refreshable table instead. This does not modify Excel until the user confirms the action.", input_schema: { type: "object", additionalProperties: false, properties: { result_id: { type: "string" }, table_name: { type: "string" } }, required: ["result_id", "table_name"] } },
] as const;

const TOOLS = [CLARIFICATION_TOOL, ...BASE_TOOLS.filter(tool => !CATALOG_TOOL_NAMES.has(tool.name)), ...CATALOG_AGENT_TOOLS];

const CATALOG_CONTEXT = (catalogs: string[]) => `SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, CASE WHEN table_type='VIEW' THEN 'view' ELSE 'table' END AS kind, '' AS description FROM information_schema.tables WHERE table_catalog IN (${catalogs.map(literal).join(",")}) AND table_schema NOT IN ('information_schema','pg_catalog') UNION ALL SELECT database_name, schema_name, function_name, CASE WHEN function_type IN ('macro','table_macro') THEN 'macro' ELSE function_type END, COALESCE(description, comment, '') FROM duckdb_functions() WHERE database_name IN (${catalogs.map(literal).join(",")}) ORDER BY 1,2,4,3`;

export const DEFAULT_MODEL = DEFAULT_AI_MODEL;
export class AgentSession {
  readonly messages: AgentMessage[] = [];
  private readonly results = new Map<string, { result: QueryResult; sql: string; connection: string }>();
  private readonly prompts = new Map<string, string>();

  reset(): void { this.messages.length = 0; this.results.clear(); this.prompts.clear(); }
  restore(messages: AgentMessage[]): void {
    this.reset();
    this.messages.push(...JSON.parse(JSON.stringify(messages)) as AgentMessage[]);
    sanitizeConversation(this.messages);
  }
  snapshot(): AgentMessage[] { return JSON.parse(JSON.stringify(this.messages)) as AgentMessage[]; }

  async run(apiKey: string, model: string, prompt: string, connection: DesktopConnection, callbacks: AgentCallbacks, signal: AbortSignal, options: AIRequestOptions = {}): Promise<void> {
    if (!apiKey.trim()) throw new Error("An Anthropic API key is required.");
    if (!prompt.trim()) throw new Error("Enter a question for the agent.");
    this.messages.push({ role: "user", content: prompt.trim() });
    const repeated = new Map<string, number>();
    const runId = crypto.randomUUID();
    let sqlFailures = 0;
    let visibleText = "";
    trace({ event: "run_start", run_id: runId, model: model || DEFAULT_MODEL, connection: connection.name, catalog: connection.catalog, prompt_chars: prompt.trim().length });

    try {
      callbacks.onStage?.("Exploring your catalog…");
      const inventoryStarted = performance.now();
      let systemPrompt: string;
      try { systemPrompt = await this.systemPrompt(connection, signal); recordAITool("list_tables", performance.now() - inventoryStarted, "complete"); }
      catch (error) { recordAITool("list_tables", performance.now() - inventoryStarted, signal.aborted ? "stopped" : "error"); throw error; }
      for (let round = 0; round < 20; round++) {
        if (signal.aborted) throw new DOMException("Stopped", "AbortError");
        sanitizeConversation(this.messages);
        trace({ event: "round_start", run_id: runId, round: round + 1, message_count: this.messages.length });
        let firstTextChunk = true;
        const roundCallbacks: AgentCallbacks = { ...callbacks, onText: (chunk) => {
          if (firstTextChunk && chunk) { firstTextChunk = false; const separator = agentRoundSeparator(visibleText, chunk); if (separator) callbacks.onText(separator); }
          callbacks.onText(chunk);
        } };
        callbacks.onStage?.("Thinking…");
        const { content: blocks } = await requestAnthropic(apiKey, model || DEFAULT_MODEL, systemPrompt, TOOLS, this.messages, roundCallbacks, signal, options);
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
          this.messages.push({ role: "user", content: results });
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
              const content = await this.execute(call.name, call.input, connection.name, { ...callbacks, onQueryResult: result => callbacks.onQueryResult?.({ ...result, toolCallId: call.id }) }, signal);
              signal.throwIfAborted();
              recordAITool(call.name, performance.now() - started, "complete");
              callbacks.onTool(call.name, "done", completedToolDetail(call.name, call.input, content), call.id);
              trace({ event: "tool_result", run_id: runId, round: round + 1, tool_use_id: call.id, tool: call.name, elapsed_ms: Math.round(performance.now() - started), result_chars: content.length });
              results.push({ type: "tool_result", tool_use_id: call.id, content });
            } catch (error) {
              recordAITool(call.name, performance.now() - started, signal.aborted || (error as Error).name === "AbortError" ? "stopped" : "error");
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
        } catch (error) {
          repairAgentHistory(this.messages);
          throw error;
        }
      }
      throw new Error("The agent reached the 20-round safety limit.");
    } catch (error) {
      trace({ event: (error as Error).name === "AbortError" ? "run_stopped" : "run_error", run_id: runId, error: error instanceof Error ? error.message : String(error) });
      throw error;
    }
  }

  private async systemPrompt(connection: DesktopConnection, signal: AbortSignal): Promise<string> {
    const cacheKey = `${connection.name}\0${(connection.catalogs ?? [connection.catalog]).join("\0")}\0${connection.authentication}`;
    const cached = this.prompts.get(cacheKey);
    if (cached) return cached;
    let objects: AgentCatalogObject[] = [];
    let inventoryError: string | undefined;
    try {
      const result = await host.query(CATALOG_CONTEXT(connection.catalogs ?? [connection.catalog]), connection.name, true, 20_000, signal);
      objects = records(result).map((row) => ({
        catalog: String(row.catalog ?? connection.catalog),
        schema: String(row.schema ?? "main"),
        name: String(row.name ?? ""),
        kind: String(row.kind ?? "object"),
        description: row.description == null ? undefined : String(row.description),
      })).filter((item) => item.name);
    } catch (error) {
      signal.throwIfAborted();
      inventoryError = error instanceof Error ? error.message : String(error);
    }
    const value = buildExcelAgentSystemPrompt({
      connection: { name: connection.name, catalog: connection.catalog, catalogs: connection.catalogs, authentication: connection.authentication },
      objects,
      inventoryError,
    }) + "\n\nFor Excel output, prefer stage_result_to_new_sheet: it stages a refreshable Power Query using the original SQL and saved Cupola connection. Loading reruns the full query; preview rows are not copied. Excel Refresh All can refresh the resulting table. Static replacement of an existing table is only for an explicit static-snapshot request. Never claim a table has been inserted before user confirmation.";
    signal.throwIfAborted();
    this.prompts.set(cacheKey, value);
    return value;
  }

  private async execute(name: string, input: Record<string, unknown>, connection: string, callbacks: AgentCallbacks, signal: AbortSignal): Promise<string> {
    if (name === "ask_clarification") {
      if (!callbacks.onClarification) throw new Error("Clarification is unavailable. Ask your question in the answer and wait for the user's next message.");
      const answer = await callbacks.onClarification(clarificationInput(input), signal);
      signal.throwIfAborted();
      return JSON.stringify({ answer });
    }
    if (name === "run_sql") {
      const result = await host.query(String(input.sql ?? ""), connection, true, 10_000, signal);
      signal.throwIfAborted();
      const id = crypto.randomUUID();
      this.results.set(id, { result, sql: String(input.sql), connection });
      callbacks.onResult(result);
      callbacks.onQueryResult?.({ id, sql: String(input.sql), result, connection, catalog: "", scope: resultScope(input.scope) });
      return JSON.stringify({ result_id: id, columns: result.columns, rows: result.rows.slice(0, 20), row_count: result.rowCount, truncated: result.rows.length > 20 || result.truncated });
    }
    if (name === "read_query_results") {
      const result = this.results.get(String(input.result_id ?? ""))?.result;
      if (!result) throw new Error("That query result is no longer available.");
      const offset = Math.max(0, Number(input.offset ?? 0));
      const limit = Math.max(1, Math.min(100, Number(input.limit ?? 20)));
      return JSON.stringify({ columns: result.columns, rows: result.rows.slice(offset, offset + limit), offset, row_count: result.rowCount });
    }
    if (CATALOG_TOOL_NAMES.has(name)) return executeCatalogTool(name, input, sql => host.query(sql, connection, true, 20_000, signal));
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
      const saved = this.results.get(resultId);
      const result = saved?.result;
      if (!result) throw new Error("That query result is no longer available. Run the query again before staging a workbook action.");
      const action: WorkbookWriteAction = {
        id: crypto.randomUUID(),
        resultId,
        mode: name === "stage_result_to_new_sheet" ? "new_sheet" : "replace_table",
        result,
        query: name === "stage_result_to_new_sheet" ? { sql: saved!.sql, connection: saved!.connection } : undefined,
        sheetName: name === "stage_result_to_new_sheet" ? excelWorksheetName(input.sheet_name) : undefined,
        tableName: String(input.table_name ?? "VGI_Result"),
      };
      callbacks.onWorkbookAction(action);
      return JSON.stringify({ staged: true, action_id: action.id, mode: action.mode, sheet_name: action.sheetName, table_name: action.tableName, row_count: result.rowCount, refreshable: action.mode === "new_sheet", message: "The workbook has not changed. The user must confirm this action in the Workbench." });
    }
    throw new Error(`Unknown agent tool: ${name}`);
  }
}

function records(result: QueryResult): Record<string, unknown>[] { return result.rows.map((row) => Object.fromEntries(result.columns.map((column, index) => [column.name, row[index]]))); }
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
