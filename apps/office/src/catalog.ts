import type { AgentToolContext, CatalogFunction, QueryBackend, QueryResult } from "@query-farm/vgi-excel-core";

export function agentContext(backend: QueryBackend, createQuery?: AgentToolContext["createQuery"]): AgentToolContext {
  return {
    backend,
    createQuery,
    async listTables() {
      return rows(
        await backend.query(`SELECT table_catalog, table_schema, table_name, table_type
          FROM information_schema.tables
          WHERE table_schema NOT IN ('information_schema', 'pg_catalog')
          ORDER BY 1, 2, 3`),
      );
    },
    async listFunctions() {
      return discoverFunctions(backend);
    },
    async describeTable(input) {
      const conditions = [
        `table_schema = ${literal(input.schema)}`,
        `table_name = ${literal(input.table)}`,
        input.catalog ? `table_catalog = ${literal(input.catalog)}` : null,
      ].filter(Boolean);
      return rows(
        await backend.query(`SELECT table_catalog, table_schema, table_name, column_name, data_type, is_nullable
          FROM information_schema.columns WHERE ${conditions.join(" AND ")} ORDER BY ordinal_position`),
      );
    },
  };
}

export async function discoverFunctions(backend: QueryBackend): Promise<CatalogFunction[]> {
  const [result, argumentResult] = await Promise.all([
    backend.query(`SELECT database_name, schema_name, function_name, function_type, parameters, parameter_types, return_type, description
      FROM duckdb_functions()
      WHERE database_name NOT IN ('system', 'temp')
      ORDER BY 1, 2, 3`),
    backend.query(`SELECT catalog_name, schema_name, function_name, arg_position, arg_name, arg_type, arg_description, is_named, is_positional, is_varargs, arg_default, arg_choices, arg_range, arg_pattern
      FROM vgi_function_arguments()
      ORDER BY 1, 2, 3, field_index`).catch((): QueryResult => ({ columns: [], rows: [], rowCount: 0 })),
  ]);
  const rich = new Map<string, CatalogFunction["parameters"]>();
  for (const row of rows(argumentResult)) {
    const key = functionKey(row.catalog_name, row.schema_name, row.function_name);
    const parameters = rich.get(key) ?? [];
    parameters.push({
      name: String(row.arg_name ?? ""),
      type: String(row.arg_type ?? ""),
      kind: truthy(row.is_named) ? "named" : truthy(row.is_positional) ? "positional" : truthy(row.is_varargs) ? "varargs" : "other",
      position: row.arg_position == null ? null : Number(row.arg_position),
      ...(row.arg_description ? { description: String(row.arg_description) } : {}),
      ...(row.arg_default != null ? { default: parseJson(row.arg_default) } : {}),
      ...(row.arg_choices != null ? { choices: parseArray(row.arg_choices) } : {}),
      ...(row.arg_range ? { range: String(row.arg_range) } : {}),
      ...(row.arg_pattern ? { pattern: String(row.arg_pattern) } : {}),
    });
    rich.set(key, parameters);
  }
  return rows(result).map((row) => ({
    catalog: String(row.database_name),
    schema: String(row.schema_name),
    name: String(row.function_name),
    kind: functionKind(String(row.function_type)),
    parameters: rich.get(functionKey(row.database_name, row.schema_name, row.function_name)) ?? parseParameters(row.parameters, row.parameter_types),
    returnType: row.return_type ? String(row.return_type) : undefined,
    description: row.description ? String(row.description) : undefined,
  }));
}

export function rows(result: QueryResult): Record<string, unknown>[] {
  return result.rows.map((row) => Object.fromEntries(result.columns.map((column, index) => [column.name, row[index]])));
}

function parseParameters(namesValue: unknown, typesValue: unknown): Array<{ name: string; type: string }> {
  const names = parseList(namesValue);
  const types = parseList(typesValue);
  return types.map((type, index) => ({ name: names[index] || `arg_${index + 1}`, type }));
}

function functionKey(catalog: unknown, schema: unknown, name: unknown): string { return `${String(catalog ?? "")}\0${String(schema ?? "")}\0${String(name ?? "")}`; }
function functionKind(value: string): CatalogFunction["kind"] {
  switch (value.toLowerCase()) {
    case "table": case "table_macro": return "table";
    case "aggregate": return "aggregate";
    case "table_in_out": return "table_in_out";
    case "table_buffering": return "buffering";
    default: return "scalar";
  }
}
function truthy(value: unknown): boolean { return value === true || value === 1 || String(value).toLowerCase() === "true"; }
function parseJson(value: unknown): unknown { if (typeof value !== "string") return value; try { return JSON.parse(value); } catch { return value; } }
function parseArray(value: unknown): unknown[] { const parsed = parseJson(value); return Array.isArray(parsed) ? parsed : []; }

function parseList(value: unknown): string[] {
  if (Array.isArray(value)) return value.map(String);
  const text = String(value ?? "").trim();
  if (!text || text === "[]") return [];
  try {
    const parsed = JSON.parse(text) as unknown;
    if (Array.isArray(parsed)) return parsed.map(String);
  } catch {
    // DuckDB renders LIST values as [value, value] in the browser Arrow path.
  }
  return text.replace(/^\[|\]$/g, "").split(",").map((item) => item.trim().replace(/^['\"]|['\"]$/g, ""));
}

function literal(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}
