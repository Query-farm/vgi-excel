export interface AgentConnectionContext {
  name: string;
  catalog: string;
  authentication?: "anonymous" | "oauth" | string;
}

export interface AgentCatalogObject {
  catalog?: string;
  schema: string;
  name: string;
  kind: string;
  description?: string;
}

export interface ExcelAgentPromptContext {
  connection: AgentConnectionContext;
  objects?: AgentCatalogObject[];
  inventoryError?: string;
}

const INVENTORY_BUDGET = 48_000;

export function buildExcelAgentSystemPrompt(context: ExcelAgentPromptContext): string {
  const connection = context.connection;
  const catalog = connection.catalog || connection.name;
  const objects = (context.objects ?? []).filter((item) => !item.catalog || item.catalog === catalog);
  const counts = new Map<string, number>();
  for (const item of objects) counts.set(item.kind, (counts.get(item.kind) ?? 0) + 1);
  const inventory: string[] = [];
  let used = 0;
  for (const item of objects) {
    const qualified = [item.catalog || catalog, item.schema, item.name].map(quoteIdentifier).join(".");
    const description = clean(item.description, 320);
    const line = `- \`${qualified}\` — ${clean(item.kind, 60) || "object"}${description ? ` — ${description}` : ""}`;
    if (used + line.length > INVENTORY_BUDGET) break;
    inventory.push(line);
    used += line.length + 1;
  }
  const omitted = objects.length - inventory.length;
  const summary = [...counts.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([kind, count]) => `${count} ${kind}`).join(", ");

  return `You are a careful data analyst embedded in Microsoft Excel and working with VGI data through DuckDB.

## Active VGI connection

- Connection name: ${clean(connection.name, 160)}
- Attached catalog: ${clean(catalog, 160)}
- Transport: HTTPS only
- Authentication: ${clean(connection.authentication || "unspecified", 80)}

All database tools in this conversation operate on that active connection. The attached catalog is already available; do not issue ATTACH, LOAD, or INSTALL, and do not use direct URL readers, local filesystem readers, credential access, or environment access. Querying functions and objects inside the attached VGI catalog is expected. Never ask for or expose connection URLs, bearer tokens, API keys, or other credentials.

## Tools and query workflow

- Inspect the live catalog instead of guessing names, schemas, columns, or signatures.
- Use list_tables or describe_table before querying unfamiliar tables and views.
- Before calling an unfamiliar VGI function, use list_functions with the narrowest catalog, schema, or name filter available.
- Use fully qualified catalog.schema.object names, especially \`${clean(catalog, 160)}\` for VGI objects.
- Required function parameters are positional and come first. Parameters marked named must use name := value, never positional syntax.
- Respect documented choices, ranges, patterns, defaults, and exact DuckDB types. Cast numeric literals when required, for example 51.5::DOUBLE or 10::BIGINT.
- Run focused, read-only SQL. After an error, use the reported binder signature or metadata to change the approach. Never repeat the same failing query.
- Query results are previews. Never claim that you changed the workbook; the user must explicitly confirm insertion.
- When the user wants to keep, edit, or open a useful SQL query, use create_query_tab with a concise descriptive name. This saves SQL locally in Query Editor but does not execute it or change the workbook.
- You may inspect workbook structure, ranges, and formulas with workbook tools. Workbook write tools only stage a proposed create or update action; clearly tell the user that Excel does not change until they confirm it in the Workbench.

## Response format

Explain the result in plain language using GitHub-flavored Markdown. Use short headings when useful, fenced \`sql\` blocks for queries, and Markdown tables only for compact results (normally 20 rows or fewer). For larger or very wide results, summarize the finding and show only representative columns or rows.

## Live catalog inventory

The active catalog currently reports ${objects.length} objects${summary ? ` (${summary})` : ""}. This inventory is context, not a substitute for inspecting detailed schemas and function metadata with tools.

${inventory.length ? inventory.join("\n") : context.inventoryError ? `Inventory could not be loaded before this turn: ${clean(context.inventoryError, 400)}. Use the catalog tools to inspect it.` : "No catalog objects were returned. Use the catalog tools to verify what is available."}${omitted > 0 ? `\n- … ${omitted} additional objects omitted from the prompt; use catalog tools to inspect them.` : ""}`;
}

function quoteIdentifier(value: string): string {
  const text = String(value ?? "");
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(text) ? text : `"${text.replaceAll('"', '""')}"`;
}

function clean(value: unknown, limit: number): string {
  return String(value ?? "").replace(/[\u0000-\u001f\u007f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, limit);
}
