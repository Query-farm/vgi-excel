const DETERMINISTIC_TOOLS = new Set(["list_tables", "list_functions", "describe_table", "read_query_results", "workbook_overview", "read_range", "list_formulas", "stage_result_to_new_sheet", "stage_result_to_table"]);
export const MAX_IDENTICAL_TOOL_CALLS = 2;

export function recordToolCall(counts: Map<string, number>, name: string, input: unknown): { block: boolean; count: number } {
  if (!DETERMINISTIC_TOOLS.has(name)) return { block: false, count: 0 };
  const key = `${name}\0${stableStringify(input)}`;
  const count = (counts.get(key) ?? 0) + 1;
  counts.set(key, count);
  return { block: count > MAX_IDENTICAL_TOOL_CALLS, count };
}

export function repeatedCallMessage(name: string, count: number): string {
  return `You already called ${name} with these exact arguments ${count - 1} times and received the same result. Use that result, answer with the information available, or explain what is missing; do not call it again.`;
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "null";
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  return `{${Object.keys(value as object).sort().map((key) => `${JSON.stringify(key)}:${stableStringify((value as Record<string, unknown>)[key])}`).join(",")}}`;
}
