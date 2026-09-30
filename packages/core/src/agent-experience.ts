import { excelWorksheetName } from "./excel-names.js";
import type { QueryResult } from "./types.js";
export interface AgentScope { dateRange?: string; filters?: string; assumptions?: string }
export interface AgentResultCard { toolCallId?: string; queryName?: string; id: string; sql: string; result: QueryResult; connection: string; catalog: string; scope: AgentScope }
export interface AgentClarification { question: string; options: string[]; answer?: string }
export const CLARIFICATION_TOOL = { name: "ask_clarification", description: "Ask the user one short clarifying question before querying when the date range, currency, business unit, or meaning of a measure materially changes the answer. Wait for their choice; free-text answers are also supported.", input_schema: { type: "object", additionalProperties: false, properties: { question: { type: "string" }, options: { type: "array", items: { type: "string" }, minItems: 2, maxItems: 5 } }, required: ["question", "options"] } } as const;
export const RESULT_SCOPE_SCHEMA = { type: "object", description: "Explain the SQL's date range, filters and assumptions in plain language. Do not invent facts not established by the query or user.", properties: { dateRange: { type: "string" }, filters: { type: "string" }, assumptions: { type: "string" } } } as const;
export function clarificationInput(input: Record<string, unknown>): AgentClarification {
  const question = typeof input.question === "string" ? input.question.trim().slice(0, 500) : "";
  const options = Array.isArray(input.options) ? [...new Set(input.options.filter((value): value is string => typeof value === "string" && !!value.trim()).map(value => value.trim().slice(0, 160)))].slice(0, 5) : [];
  if (!question || options.length < 2) throw new Error("A question and two to five distinct choices are required.");
  return { question, options };
}
export function resultScope(value: unknown): AgentScope {
  if (!value || typeof value !== "object") return {};
  return Object.fromEntries(["dateRange", "filters", "assumptions"].flatMap(key => typeof (value as Record<string, unknown>)[key] === "string" ? [[key, String((value as Record<string, unknown>)[key]).slice(0, 1000)]] : []));
}
export function toolStage(name: string): string {
  if (name === "ask_clarification") return "Waiting for your answer…";
  if (name === "run_sql") return "Running query…";
  if (name === "read_query_results") return "Reviewing query results…";
  if (["workbook_overview", "read_range", "list_formulas"].includes(name)) return "Reading workbook context…";
  if (name.startsWith("stage_") || name === "create_query_tab") return "Preparing your result…";
  return "Exploring your catalog…";
}

export type AgentTranscriptPart = { type: "text"; text: string } | { type: "thinking"; text: string } | { type: "tool"; id: string };
export function appendTranscript(parts: AgentTranscriptPart[] | undefined, part: AgentTranscriptPart): AgentTranscriptPart[] {
  const values = [...(parts ?? [])], last = values[values.length - 1];
  if (part.type === "tool") {
    if (!values.some(value => value.type === "tool" && value.id === part.id)) values.push(part);
  } else if (last && last.type === part.type) values[values.length - 1] = { ...last, text: last.text + part.text };
  else values.push(part);
  return values;
}
export function savedTranscript(value: unknown): AgentTranscriptPart[] | undefined {
  if (!Array.isArray(value)) return undefined;
  return value.slice(0, 500).flatMap((part): AgentTranscriptPart[] => {
    if (part?.type === "tool" && typeof part.id === "string") return [{ type: "tool", id: part.id.slice(0, 256) }];
    if ((part?.type === "text" || part?.type === "thinking") && typeof part.text === "string") return [{ type: part.type, text: part.text.slice(0, 50000) }];
    return [];
  });
}

/** A short Excel-safe title, including for results from older conversations. */
export function agentQueryName(value: unknown): string {
  return excelWorksheetName(typeof value === "string" ? value : "", "AI query");
}
