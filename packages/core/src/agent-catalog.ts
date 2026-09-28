import type { QueryResult } from "./types.js";
import { normalizeTags, filterTagsForAI, filterTagsForAIDetail, examplesForAI, parseRequiredFilters, parseCategories } from "./vgi-tags.js";

export const REQUIRED_FILTERS_RULE = "Required filters are an AND of OR-groups: the WHERE clause must constrain at least one column from every group. [[\"a\"],[\"b\",\"c\"]] requires a AND (b OR c). Queries without these filters fail at bind time; ask the user for missing filter values.";
export const CATALOG_TOOL_NAMES = new Set(["list_tables", "list_functions", "list_categories", "describe_table", "describe_function"]);
const props = { catalog: { type: "string" }, schema: { type: "string" }, query: { type: "string" }, category: { type: "string" }, cursor: { type: "string" }, limit: { type: "integer", minimum: 1, maximum: 200 } };
export const CATALOG_AGENT_TOOLS = [
  { name: "list_tables", description: "Search and page through tables and views. Includes bounded discovery metadata; use describe_table for required filters, examples, and columns.", input_schema: { type: "object", properties: props } },
  { name: "list_functions", description: "Search and page through functions and macros, including named/positional arguments, types, constraints, examples, and metadata.", input_schema: { type: "object", properties: { ...props, name: { type: "string" } } } },
  { name: "list_categories", description: "Read the controlled category registry for the selected catalog's schemas.", input_schema: { type: "object", properties: { catalog: props.catalog, schema: props.schema } } },
  { name: "describe_table", description: "Describe a table or view: columns, constraints, required WHERE-filter groups, examples and detailed VGI documentation. Inspect this before querying unfamiliar tables.", input_schema: { type: "object", properties: { catalog: props.catalog, schema: props.schema, table: { type: "string" } }, required: ["schema", "table"] } },
  { name: "describe_function", description: "Describe an exact function or macro, its named/positional arguments, defaults, choices, ranges, patterns, return type, examples, and VGI metadata.", input_schema: { type: "object", properties: { catalog: props.catalog, schema: props.schema, function: { type: "string" } }, required: ["schema", "function"] } },
] as const;
export type CatalogQuery = (sql: string) => Promise<QueryResult>;
export function catalogRecords(result: QueryResult): Record<string, unknown>[] { return result.rows.map(row => Object.fromEntries(result.columns.map((col, i) => [col.name, row[i]]))); }
const literal = (value: unknown) => `'${String(value ?? "").replaceAll("'", "''")}'`;
function parsed(value: unknown): any { if (typeof value !== "string") return value; try { return JSON.parse(value); } catch { return value; } }
function tagsOf(value: unknown): Record<string, string> { return normalizeTags(parsed(value)); }
function metadata(row: Record<string, unknown>, detail = false): Record<string, unknown> {
  const tags = tagsOf(row.tags);
  const native = parsed(row.examples);
  return {
    ...row, ...(row.comment != null ? { comment: String(row.comment).slice(0, detail ? 4000 : 500) } : {}), description: String(row.description ?? row.comment ?? "").slice(0, detail ? 4000 : 500),
    tags: detail ? filterTagsForAIDetail(tags) : filterTagsForAI(tags),
    ...(detail ? Object.fromEntries(["semantic_catalog", "semantic_entity", "semantic_member", "semantic_members", "semantic_relationships"].filter(key => tags[`vgi.${key}`]).map(key => [key, parsed(tags[`vgi.${key}`].slice(0, 4000))])) : {}),
    ...(detail ? { examples: examplesForAI(tags, Array.isArray(native) ? native.map(v => typeof v === "string" ? { sql: v } : v) : []) } : { examples: undefined }),
  };
}
function page(values: Record<string, unknown>[], input: Record<string, unknown>, truncated?: boolean): object {
  const query = String(input.query ?? input.name ?? "").toLowerCase();
  const filtered = values.filter(row => (!query || JSON.stringify(row).toLowerCase().includes(query)) && (!input.category || String((row.tags as Record<string, unknown> | undefined)?.["vgi.category"] ?? "") === input.category));
  const rawOffset = Number(input.cursor ?? 0), rawLimit = Number(input.limit ?? 100);
  const offset = Number.isInteger(rawOffset) && rawOffset >= 0 ? rawOffset : 0;
  const limit = Number.isFinite(rawLimit) ? Math.max(1, Math.min(200, Math.floor(rawLimit))) : 100;
  return { objects: filtered.slice(offset, offset + limit), total: filtered.length, next_cursor: offset + limit < filtered.length ? String(offset + limit) : null, truncated: !!truncated };
}
export async function executeCatalogTool(name: string, input: Record<string, unknown>, query: CatalogQuery): Promise<string> {
  const baseFilters = (catalogColumn: string) => [input.catalog ? `${catalogColumn}=${literal(input.catalog)}` : `${catalogColumn} NOT IN ('system','temp')`, input.schema ? `schema_name=${literal(input.schema)}` : "schema_name NOT IN ('information_schema','pg_catalog')"];
  const filters = baseFilters("database_name");
  if (name === "list_categories") {
    const result = await query(`SELECT database_name AS catalog, schema_name AS schema, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_schemas() WHERE ${filters.join(" AND ")} ORDER BY 1,2`);
    return JSON.stringify({ schemas: catalogRecords(result).map(row => ({ catalog: row.catalog, schema: row.schema, categories: parseCategories(tagsOf(row.tags)) })) });
  }
  if (name === "list_functions" || name === "describe_function") {
    if (name === "list_functions" && input.name) filters.push(`function_name ILIKE ${literal(`%${String(input.name)}%`)}`);
    if (name === "describe_function") filters.push(`function_name=${literal(input.function)}`);
    const functions = await query(`SELECT database_name AS catalog, schema_name AS schema, function_name AS name, function_type AS kind, description, return_type, CAST(to_json(parameters) AS VARCHAR) AS parameters, CAST(to_json(parameter_types) AS VARCHAR) AS parameter_types, CAST(to_json(examples) AS VARCHAR) AS examples, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_functions() WHERE ${filters.join(" AND ")} ORDER BY 1,2,3`);
    const richFilters = baseFilters("catalog_name");
    if (name === "describe_function") richFilters.push(`function_name=${literal(input.function)}`);
    if (name === "list_functions" && input.name) richFilters.push(`function_name ILIKE ${literal(`%${String(input.name)}%`)}`);
    let rich: Record<string, unknown>[] = [], warning: string | undefined;
    try { rich = catalogRecords(await query(`SELECT catalog_name, schema_name, function_name, arg_name, arg_type, arg_position, is_named, is_positional, is_varargs, arg_description, arg_default, arg_choices, arg_range, arg_pattern FROM vgi_function_arguments() WHERE ${richFilters.join(" AND ")} ORDER BY catalog_name,schema_name,function_name,field_index`)); }
    catch (error) { if ((error as Error).name === "AbortError") throw error; warning = "Rich argument metadata is unavailable. Do not infer named versus positional arguments from the fallback signature."; }
    const values = catalogRecords(functions).map(row => {
      const args = rich.filter(arg => arg.catalog_name === row.catalog && arg.schema_name === row.schema && arg.function_name === row.name);
      const names = parsed(row.parameters), types = parsed(row.parameter_types);
      const { parameters: _names, parameter_types: _types, ...value } = metadata(row, true);
      return { ...value, qualified_name: [row.catalog, row.schema, row.name].join("."), arguments: args.length ? args.map(arg => ({ name: arg.arg_name, type: arg.arg_type, position: arg.arg_position, kind: isTrue(arg.is_named) ? "named" : isTrue(arg.is_positional) ? "positional" : isTrue(arg.is_varargs) ? "varargs" : "unknown", description: arg.arg_description, default: parsed(arg.arg_default), choices: parsed(arg.arg_choices), range: arg.arg_range, pattern: arg.arg_pattern })) : (Array.isArray(types) ? types : []).map((type, i) => ({ name: names?.[i] ?? `arg_${i + 1}`, type, kind: "unknown" })) };
    });
    return JSON.stringify({ ...page(values, name === "describe_function" ? {} : input, functions.truncated), guidance: "Use positional arguments in position order; named arguments use name := value. Respect exact types, defaults, choices, ranges and patterns.", warning });
  }
  const viewFilters = baseFilters("database_name");
  if (name === "describe_table") viewFilters.push(`view_name=${literal(input.table)}`);
  if (name === "describe_table") filters.push(`table_name=${literal(input.table)}`);
  const objects = await query(`SELECT database_name AS catalog, schema_name AS schema, table_name AS name, 'table' AS kind, comment, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_tables() WHERE ${filters.join(" AND ")} UNION ALL SELECT database_name AS catalog, schema_name AS schema, view_name AS name, 'view' AS kind, comment, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_views() WHERE ${viewFilters.join(" AND ")} ORDER BY 1,2,3`);
  if (name === "list_tables") return JSON.stringify(page(catalogRecords(objects).map(row => metadata(row)), input, objects.truncated));
  if (name !== "describe_table") throw new Error(`Unknown catalog tool: ${name}`);
  const columns = await query(`SELECT database_name AS catalog, schema_name AS schema, table_name AS name, column_name, data_type, is_nullable, comment, column_default FROM duckdb_columns() WHERE ${filters.join(" AND ")} ORDER BY database_name,schema_name,table_name,column_index`);
  const constraints = await query(`SELECT database_name AS catalog, schema_name AS schema, table_name AS name, constraint_type, constraint_text, CAST(to_json(constraint_column_names) AS VARCHAR) AS columns FROM duckdb_constraints() WHERE ${filters.join(" AND ")} ORDER BY constraint_index`);
  return JSON.stringify({ objects: catalogRecords(objects).map(row => {
    const match = (other: Record<string, unknown>) => other.catalog === row.catalog && other.schema === row.schema && other.name === row.name;
    const required = parseRequiredFilters(tagsOf(row.tags));
    return { ...metadata(row, true), columns: catalogRecords(columns).filter(match), constraints: catalogRecords(constraints).filter(match), required_filters: required, ...(required.length ? { required_filters_rule: REQUIRED_FILTERS_RULE } : {}) };
  }) });
}
function isTrue(value: unknown): boolean { return value === true || value === 1 || value === "true"; }
