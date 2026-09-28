export interface AIUsage { inputTokens: number; outputTokens: number; cacheReadTokens: number; cacheWriteTokens: number }
export interface AIRequestDiagnostic extends AIUsage { elapsedMs: number; firstTextMs?: number; outcome: "complete" | "stopped" | "error" }
const requests: AIRequestDiagnostic[] = [];
const number = (value: unknown): number => typeof value === "number" && Number.isFinite(value) && value >= 0 ? Math.floor(value) : 0;
export function aiUsage(value: Record<string, unknown> = {}): AIUsage { return { inputTokens: number(value.input_tokens), outputTokens: number(value.output_tokens), cacheReadTokens: number(value.cache_read_input_tokens), cacheWriteTokens: number(value.cache_creation_input_tokens) }; }
export function recordAIRequest(value: AIRequestDiagnostic): void {
  // Explicit numeric allowlist: never retain prompts, SQL, names, keys or provider objects.
  requests.push({ elapsedMs: number(value.elapsedMs), ...(value.firstTextMs === undefined ? {} : { firstTextMs: number(value.firstTextMs) }), outcome: value.outcome === "complete" || value.outcome === "stopped" ? value.outcome : "error", inputTokens: number(value.inputTokens), outputTokens: number(value.outputTokens), cacheReadTokens: number(value.cacheReadTokens), cacheWriteTokens: number(value.cacheWriteTokens) });
  if (requests.length > 50) requests.shift();
}
export function readAIRequestDiagnostics(): AIRequestDiagnostic[] { return requests.map(value => ({ ...value })); }
export function clearAIRequestDiagnostics(): void { requests.length = 0; toolCalls.length = 0; }
export interface AIToolDiagnostic { kind: "query" | "catalog" | "clarification" | "workbook" | "editor"; elapsedMs: number; outcome: "complete" | "stopped" | "error" }
const toolCalls: AIToolDiagnostic[] = [];
export function recordAITool(name: string, elapsedMs: number, outcome: AIToolDiagnostic["outcome"]): void {
  const kind = name === "run_sql" ? "query" : name === "ask_clarification" ? "clarification" : name === "create_query_tab" ? "editor" : ["workbook_overview", "read_range", "list_formulas", "stage_result_to_new_sheet", "stage_result_to_table"].includes(name) ? "workbook" : "catalog";
  toolCalls.push({ kind, elapsedMs: number(elapsedMs), outcome: outcome === "complete" || outcome === "stopped" ? outcome : "error" });
  if (toolCalls.length > 100) toolCalls.shift();
}
export function readAIToolDiagnostics(): AIToolDiagnostic[] { return toolCalls.map(value => ({ ...value })); }
