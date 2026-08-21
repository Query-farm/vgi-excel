using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace QueryFarm.Vgi.ExcelDna;

internal static class AgentSqlPolicy
{
    private static readonly string[] ReadPrefixes = { "SELECT", "WITH", "DESCRIBE", "DESC", "SHOW", "EXPLAIN" };

    public static void AssertReadOnly(string sql)
    {
        var masked = Mask(sql).Trim();
        var keyword = Regex.Match(masked, @"^([A-Za-z]+)").Groups[1].Value.ToUpperInvariant();
        if (!ReadPrefixes.Contains(keyword))
            throw new InvalidOperationException("The AI agent may only run read-only SQL statements.");
        var withoutTrailing = masked.EndsWith(";") ? masked.Substring(0, masked.Length - 1) : masked;
        if (withoutTrailing.Contains(";"))
            throw new InvalidOperationException("The AI agent may run only one SQL statement at a time.");
        if (Regex.IsMatch(masked, @"\b(ATTACH|DETACH|INSTALL|LOAD|COPY|EXPORT|IMPORT|CALL|PRAGMA|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP|TRUNCATE|REPLACE|GRANT|REVOKE|VACUUM)\b", RegexOptions.IgnoreCase))
            throw new InvalidOperationException("The AI agent may not change data, connections, extensions, or external state.");
        if (Regex.IsMatch(masked, @"\b(read_csv(?:_auto)?|read_json(?:_auto)?|read_ndjson(?:_auto)?|read_parquet|parquet_scan|read_blob|read_text|read_xlsx|glob|sqlite_scan|postgres_scan|mysql_scan|delta_scan|iceberg_scan|query|query_table|getenv|duckdb_secrets)\s*\(", RegexOptions.IgnoreCase))
            throw new InvalidOperationException("The AI agent may not read local files, arbitrary URLs, secrets, or external databases.");
        if (Regex.IsMatch(sql, "(?i)['\"]\\s*(?:https?|file)://|\\bFROM\\s+['\"]|[A-Za-z]:\\\\"))
            throw new InvalidOperationException("The AI agent may query only the selected VGI catalog and safe database metadata.");
    }

    private static string Mask(string sql)
    {
        var noComments = Regex.Replace(sql, @"--[^\r\n]*|/\*[\s\S]*?\*/", " ");
        return Regex.Replace(noComments, "'(?:''|[^'])*'|\"(?:\"\"|[^\"])*\"", " ");
    }
}
