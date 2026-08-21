using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace QueryFarm.Vgi.ExcelDna;

/// <summary>
/// Creates an ordinary Excel Power Query whose data source is the Cupola ODBC
/// driver. The workbook stores only the Cupola connection name and SQL. The
/// driver resolves endpoint, ATTACH options, and credentials from Cupola's
/// per-user connection store.
/// </summary>
internal static class PowerQueryBridge
{
    internal const string DriverEnvironmentVariable = "CUPOLA_ODBC_DRIVER_NAME";
    internal const string DefaultDriverName = "Cupola for Excel";

    internal static string DriverName()
    {
        var configured = Environment.GetEnvironmentVariable(DriverEnvironmentVariable);
        return string.IsNullOrWhiteSpace(configured) ? DefaultDriverName : configured.Trim();
    }

    internal static string ConnectionString(string connectionName, string? driverName = null)
    {
        if (string.IsNullOrWhiteSpace(connectionName)) throw new ArgumentException("A Cupola connection name is required.");
        return $"Driver={{{EscapeOdbc(driverName ?? DriverName())}}};CupolaConnection={{{EscapeOdbc(connectionName.Trim())}}};";
    }

    internal static string Formula(string sql, string connectionName, string? driverName = null)
    {
        if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required.");
        AgentSqlPolicy.AssertReadOnly(sql);
        return "let\n" +
               $"    Source = Odbc.Query(\"{EscapeM(ConnectionString(connectionName, driverName))}\", \"{EscapeM(sql.Trim())}\")\n" +
               "in\n" +
               "    Source";
    }

    internal static object Create(string sql, string connectionName, string? requestedName, bool loadToWorksheet)
    {
        var connection = ConnectionStore.Resolve(connectionName);
        dynamic app = global::ExcelDna.Integration.ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        var queryName = UniqueQueryName(book, requestedName);
        dynamic query = book.Queries.Add(queryName, Formula(sql, connection.Name),
            $"Created by {ProductInfo.Name} {ProductInfo.Version} from connection {connection.Name}.");

        if (!loadToWorksheet)
            return new { query = queryName, loaded = false, message = "The query was added to Queries & Connections. Use Load To… to place it in the workbook." };

        if (!IsDriverRegistered(DriverName()))
            return new { query = queryName, loaded = false, message = $"The query was added to Queries & Connections. Install the “{DriverName()}” ODBC driver, then choose Load To… or Refresh." };

        try
        {
            dynamic sheet = book.Worksheets.Add(After: book.Worksheets[book.Worksheets.Count]);
            sheet.Name = UniqueSheetName(book, queryName);
            var source = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName};Extended Properties=\"\"";
            dynamic table = sheet.ListObjects.Add(0, source, Type.Missing, 1, sheet.Range["A1"]);
            table.Name = UniqueTableName(book, queryName);
            table.QueryTable.CommandType = 2;
            table.QueryTable.CommandText = new[] { $"SELECT * FROM [{queryName}]" };
            table.QueryTable.BackgroundQuery = true;
            table.QueryTable.Refresh(false);
            return new { query = queryName, loaded = true, sheet = Convert.ToString(sheet.Name), table = Convert.ToString(table.Name), message = "Power Query created and refresh started." };
        }
        catch (Exception error)
        {
            // The WorkbookQuery remains useful and visible even when the ODBC
            // driver is not installed yet or Excel cannot load it immediately.
            return new { query = queryName, loaded = false, message = "The query was created, but Excel could not load it to a worksheet: " + error.Message };
        }
    }

    private static string UniqueQueryName(dynamic book, string? requested)
    {
        var root = CleanName(requested, "Cupola Query", 80);
        var names = new List<string>();
        foreach (dynamic query in book.Queries) names.Add(Convert.ToString(query.Name) ?? "");
        return Unique(root, names.ToArray(), 80);
    }

    internal static bool IsDriverRegistered(string driverName)
    {
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(hive, view);
                using var drivers = root.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers");
                if (drivers?.GetValue(driverName) is string status && status.Equals("Installed", StringComparison.OrdinalIgnoreCase)) return true;
                using var driver = root.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\" + driverName);
                if (driver is not null) return true;
            }
            catch { }
        }
        return false;
    }

    private static string UniqueSheetName(dynamic book, string requested)
    {
        var root = WorkbookBridge.NormalizeWorksheetName(string.IsNullOrWhiteSpace(requested) ? "Cupola Data" : requested);
        var names = new List<string>();
        foreach (dynamic sheet in book.Worksheets) names.Add(Convert.ToString(sheet.Name) ?? "");
        return Unique(root, names.ToArray(), 31);
    }

    private static string UniqueTableName(dynamic book, string requested)
    {
        var root = System.Text.RegularExpressions.Regex.Replace(requested, @"[^A-Za-z0-9_]", "_");
        if (string.IsNullOrWhiteSpace(root) || char.IsDigit(root[0])) root = "Cupola_" + root;
        var names = new List<string>();
        foreach (dynamic sheet in book.Worksheets)
        foreach (dynamic table in sheet.ListObjects)
            names.Add(Convert.ToString(table.Name) ?? "");
        return Unique(root, names.ToArray(), 200);
    }

    private static string CleanName(string? value, string fallback, int maximum)
    {
        var source = string.IsNullOrWhiteSpace(value) ? fallback : value!.Trim();
        var cleaned = new string(source
            .Where(character => "\\/?*[]:".IndexOf(character) < 0).ToArray());
        return cleaned.Substring(0, Math.Min(maximum, cleaned.Length));
    }

    private static string Unique(string root, string[] existing, int maximum)
    {
        if (!existing.Any(name => string.Equals(name, root, StringComparison.OrdinalIgnoreCase))) return root;
        for (var index = 2; ; index++)
        {
            var suffix = $" ({index})";
            var candidate = root.Substring(0, Math.Min(root.Length, maximum - suffix.Length)) + suffix;
            if (!existing.Any(name => string.Equals(name, candidate, StringComparison.OrdinalIgnoreCase))) return candidate;
        }
    }

    private static string EscapeM(string value) => value.Replace("\"", "\"\"");
    private static string EscapeOdbc(string value) => value.Replace("}", "}}");
}
