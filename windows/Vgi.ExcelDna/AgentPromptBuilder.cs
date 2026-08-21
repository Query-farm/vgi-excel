using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace QueryFarm.Vgi.ExcelDna;

internal static class AgentPromptBuilder
{
    private const int InventoryBudget = 48_000;

    public static string Build(VgiConnection connection, QueryResult? inventory = null, string? inventoryError = null)
    {
        var objects = inventory is null ? Array.Empty<CatalogObject>() : Objects(inventory, connection.Catalog).ToArray();
        var counts = string.Join(", ", objects.GroupBy(item => item.Kind).OrderBy(group => group.Key).Select(group => $"{group.Count()} {group.Key}"));
        var lines = new List<string>();
        var used = 0;
        foreach (var item in objects)
        {
            var line = $"- `{Identifier(item.Catalog)}.{Identifier(item.Schema)}.{Identifier(item.Name)}` — {Clean(item.Kind, 60)}";
            if (!string.IsNullOrWhiteSpace(item.Description)) line += " — " + Clean(item.Description, 320);
            if (used + line.Length > InventoryBudget) break;
            lines.Add(line);
            used += line.Length + 1;
        }
        var omitted = objects.Length - lines.Count;
        var context = lines.Count > 0
            ? string.Join("\n", lines)
            : !string.IsNullOrWhiteSpace(inventoryError)
                ? $"Inventory could not be loaded before this turn: {Clean(inventoryError, 400)}. Use the catalog tools to inspect it."
                : "No catalog objects were returned. Use the catalog tools to verify what is available.";
        if (omitted > 0) context += $"\n- … {omitted} additional objects omitted from the prompt; use catalog tools to inspect them.";

        return $@"You are a careful data analyst embedded in Microsoft Excel and working with VGI data through DuckDB.

## Active VGI connection

- Connection name: {Clean(connection.Name, 160)}
- Attached catalog: {Clean(connection.Catalog, 160)}
- Transport: HTTPS only
- Authentication: {Clean(connection.Authentication, 80)}

All database tools in this conversation operate on that active connection. The attached catalog is already available; do not issue ATTACH, LOAD, or INSTALL, and do not use direct URL readers, local filesystem readers, credential access, or environment access. Querying functions and objects inside the attached VGI catalog is expected. Never ask for or expose connection URLs, bearer tokens, API keys, or other credentials.

## Tools and query workflow

- Inspect the live catalog instead of guessing names, schemas, columns, or signatures.
- Use list_tables or describe_table before querying unfamiliar tables and views.
- Before calling an unfamiliar VGI function, use list_functions.
- Use fully qualified catalog.schema.object names, especially `{Clean(connection.Catalog, 160)}` for VGI objects.
- Required function parameters are positional and come first. Parameters marked named must use name := value, never positional syntax.
- Respect documented choices, ranges, patterns, defaults, and exact DuckDB types. Cast numeric literals when required.
- Run focused, read-only SQL. After an error, inspect metadata and change the approach; never repeat the same failing query.
- Query results are previews. Never claim that you changed the workbook; the user must explicitly confirm insertion.

## Response format

Explain results in plain language using GitHub-flavored Markdown. Use short headings, fenced `sql` blocks for queries, and Markdown tables only for compact results. Summarize large or wide results.

## Live catalog inventory

The active catalog currently reports {objects.Length} objects{(counts.Length > 0 ? $" ({counts})" : "")}.

{context}";
    }

    private static IEnumerable<CatalogObject> Objects(QueryResult result, string defaultCatalog)
    {
        var positions = result.Columns.Select((column, index) => (column.Name, index)).ToDictionary(item => item.Name, item => item.index, StringComparer.OrdinalIgnoreCase);
        foreach (var row in result.Rows)
        {
            string Value(string name, string fallback = "") => positions.TryGetValue(name, out var index) && index < row.Length ? Convert.ToString(row[index]) ?? fallback : fallback;
            var name = Value("name");
            if (name.Length == 0) continue;
            yield return new CatalogObject(Value("catalog", defaultCatalog), Value("schema", "main"), name, Value("kind", "object"), Value("description"));
        }
    }

    private static string Identifier(string? value)
    {
        var text = value ?? "";
        return Regex.IsMatch(text, @"^[A-Za-z_][A-Za-z0-9_]*$") ? text : $"\"{text.Replace("\"", "\"\"")}\"";
    }
    private static string Clean(string? value, int limit)
    {
        var text = Regex.Replace(value ?? "", @"[\x00-\x1f\x7f]+", " ");
        text = Regex.Replace(text, @"\s+", " ").Trim();
        return text.Length <= limit ? text : text.Substring(0, limit);
    }

    private sealed class CatalogObject
    {
        public CatalogObject(string catalog, string schema, string name, string kind, string description)
        {
            Catalog = catalog; Schema = schema; Name = name; Kind = kind; Description = description;
        }
        public string Catalog { get; }
        public string Schema { get; }
        public string Name { get; }
        public string Kind { get; }
        public string Description { get; }
    }
}
