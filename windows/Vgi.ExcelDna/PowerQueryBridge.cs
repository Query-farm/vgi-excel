using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data.Odbc;
using System.Data.OleDb;
using Newtonsoft.Json.Linq;
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

    internal static object Create(string sql, string connectionName, string? requestedName, bool loadToWorksheet, string? requestedSheetName = null, string? requestedTableName = null)
    {
        dynamic app = global::ExcelDna.Integration.ExcelDnaUtil.Application;
        return CreateInWorkbook(app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active."), sql, connectionName, requestedName, loadToWorksheet, requestedSheetName, requestedTableName);
    }

    internal static object CreateInWorkbook(object workbook, string sql, string connectionName, string? requestedName, bool loadToWorksheet, string? requestedSheetName = null, string? requestedTableName = null)
    {
        var connection = ResolveConnection(connectionName);
        var driverName = DriverName();
        RequireDriver(driverName, InstalledDrivers());
        var formula = Formula(sql, connection.Name, driverName);
        ValidateOdbcConnection(connection, driverName);
        dynamic book = workbook;
        var queryName = UniqueQueryName(book, requestedName);
        dynamic query = book.Queries.Add(queryName, formula,
            $"Created by {ProductInfo.Name} {ProductInfo.Version} from connection {connection.Name}.");

        if (!loadToWorksheet)
            return new { query = queryName, loaded = false, message = "The query was added to Queries & Connections. Use Load To… to place it in the workbook." };

        try
        {
            dynamic sheet = book.Worksheets.Add(After: book.Worksheets[book.Worksheets.Count]);
            sheet.Name = UniqueSheetName(book, requestedSheetName ?? queryName);
            var source = MashupSource(queryName);
            dynamic table = sheet.ListObjects.Add(0, source, Type.Missing, 1, sheet.Range["A1"]);
            table.Name = UniqueTableName(book, requestedTableName ?? queryName);
            table.QueryTable.CommandType = 2;
            table.QueryTable.CommandText = new[] { $"SELECT * FROM [{queryName}]" };
            table.QueryTable.BackgroundQuery = true;
            table.QueryTable.Refresh(true);
            return new { query = queryName, loaded = true, sheet = Convert.ToString(sheet.Name), table = Convert.ToString(table.Name), message = "Power Query created and refresh started. If Excel asks for ODBC authentication, choose Default or Custom and leave credentials blank." };
        }
        catch (Exception error)
        {
            // The WorkbookQuery remains available if Excel requires first-use
            // authentication setup or cannot load the worksheet immediately.
            return new { query = queryName, loaded = false, message = "The query was created, but Excel could not load it to a worksheet. In Data → Queries & Connections, refresh it; if prompted, choose Default or Custom with no extra credentials. " + error.Message };
        }
    }

    private static string UniqueQueryName(dynamic book, string? requested)
    {
        var root = CleanName(requested, "Cupola Query", 80);
        var names = new List<string>();
        foreach (dynamic query in book.Queries) names.Add(Convert.ToString(query.Name) ?? "");
        return Unique(root, names.ToArray(), 80);
    }

    internal static string MashupSource(string queryName)
    {
        var source = new OleDbConnectionStringBuilder { Provider = "Microsoft.Mashup.OleDb.1", DataSource = "$Workbook$" };
        source["Location"] = queryName;
        source["Extended Properties"] = "";
        return "OLEDB;" + source.ConnectionString;
    }

    internal static void ValidateOdbcConnection(VgiConnection expected, string driverName)
    {
        try
        {
            using var source = new OdbcConnection(ConnectionString(expected.Name, driverName));
            source.Open();
            using var command = source.CreateCommand();
            command.CommandTimeout = 30;
            var attachments = ConnectionStore.ResolveAttachments(expected);
            command.CommandText = "SELECT contract_version, connection_name, catalog_alias, location, authentication, attach_options" + (expected.IsProfile ? ", member_name" : "") + " FROM cupola_connection_info()";
            var identities = new List<JObject>();
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var identity = new JObject();
                    for (var index = 0; index < reader.FieldCount; index++) identity[reader.GetName(index)] = JToken.FromObject(reader.GetValue(index));
                    identities.Add(identity);
                }
            }
            ValidateIdentities(expected, attachments, identities);
            command.CommandText = "SELECT count(*) FROM duckdb_databases() WHERE database_name = ? AND type = 'vgi'";
            foreach (var member in attachments)
            {
                command.Parameters.Clear();
                command.Parameters.AddWithValue("catalog", member.Catalog);
                if (Convert.ToInt32(command.ExecuteScalar()) != 1) throw new InvalidOperationException("An expected VGI catalog is not attached.");
            }

        }
        catch (Exception error) when (error is OdbcException || error is InvalidOperationException || error is ArgumentException || error is Newtonsoft.Json.JsonException || error is FormatException || error is OverflowException)
        {
            // Do not propagate ODBC diagnostics: third-party drivers can echo SQL and connection material.
            throw new InvalidOperationException("Power Query connection validation failed. The driver must support CupolaConnection and attach the selected HTTPS catalog with matching settings. Check the driver installation and sign-in in Cupola. No workbook query was created.");
        }
    }

    internal static void ValidateIdentities(VgiConnection profile, IReadOnlyList<VgiConnection> attachments, IReadOnlyList<JObject> identities)
    {
        if (identities.Count != attachments.Count) throw new InvalidOperationException("The ODBC attachment count does not match the selected connection.");
        foreach (var member in attachments)
        {
            var matches = identities.Where(row => string.Equals(row.Value<string>("catalog_alias"), member.Catalog, StringComparison.Ordinal)).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The ODBC attachment identity is missing or ambiguous.");
            var identity = (JObject)matches[0].DeepClone();
            if (profile.IsProfile)
            {
                if (identity.Value<int?>("contract_version") != 2 || !string.Equals(identity.Value<string>("connection_name"), profile.Name, StringComparison.OrdinalIgnoreCase) || !string.Equals(identity.Value<string>("member_name"), member.Name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The ODBC profile identity does not match the selected connection.");
                identity["contract_version"] = 1;
                identity["connection_name"] = member.Name;
            }
            ValidateIdentity(member, identity);
        }
    }

    internal static void ValidateIdentity(VgiConnection expected, JObject identity)
    {
        if (identity.Value<int?>("contract_version") != 1 ||
            !string.Equals(identity.Value<string>("connection_name"), expected.Name, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(identity.Value<string>("catalog_alias"), expected.Catalog, StringComparison.Ordinal) ||
            !string.Equals(identity.Value<string>("location"), expected.Location, StringComparison.Ordinal) ||
            !string.Equals(identity.Value<string>("authentication"), expected.Authentication, StringComparison.Ordinal) ||
            !JToken.DeepEquals(JObject.Parse(identity.Value<string>("attach_options") ?? "null"), JObject.FromObject(expected.AttachOptions ?? new Dictionary<string, object?>())))
            throw new InvalidOperationException("The ODBC session does not match the selected Cupola connection.");
    }

    internal static VgiConnection ResolveConnection(string name)
    {
        var connection = ConnectionStore.List().FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("The selected Cupola connection no longer exists. Select a connection in Cupola and run the query again.");
        ConnectionStore.Validate(connection);
        return connection;
    }

    internal static void RequireDriver(string driverName, IEnumerable<string> installedDrivers)
    {
        if (installedDrivers.Contains(driverName, StringComparer.OrdinalIgnoreCase)) return;
        throw new InvalidOperationException($"Power Query requires the {IntPtr.Size * 8}-bit “{driverName}” ODBC driver for this Excel installation. It is not installed. No workbook query was created. A standard Haybarn/DuckDB driver or an existing DSN does not resolve Cupola connections. Install a Cupola-compatible driver; use Insert snapshot in the meantime.");
    }

    internal static bool IsDriverRegistered(string driverName) => InstalledDrivers().Contains(driverName, StringComparer.OrdinalIgnoreCase);

    internal static string[] InstalledDrivers()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // A driver registered only for the opposite bitness cannot be loaded by Excel.
        var view = IntPtr.Size == 8 ? RegistryView.Registry64 : RegistryView.Registry32;
        foreach (var hive in new[] { RegistryHive.LocalMachine })
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var registrations = root.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI");
            if (registrations is null) continue;
            foreach (var name in registrations.GetSubKeyNames())
            {
                using var driver = registrations.OpenSubKey(name);
                if (driver?.GetValue("Driver") is string path && File.Exists(Environment.ExpandEnvironmentVariables(path))) names.Add(name);
            }
        }
        return names.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
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
        var names = new List<string>();
        foreach (dynamic sheet in book.Worksheets)
        foreach (dynamic table in sheet.ListObjects)
            names.Add(Convert.ToString(table.Name) ?? "");
        return TableName(requested, names);
    }

    internal static string TableName(string? requested, IEnumerable<string> existing)
    {
        var root = System.Text.RegularExpressions.Regex.Replace((requested ?? "Cupola_Data").Trim(), @"[^A-Za-z0-9_]", "_");
        if (string.IsNullOrWhiteSpace(root) || char.IsDigit(root[0]) ||
            System.Text.RegularExpressions.Regex.IsMatch(root, @"^(?:[RC]|[A-Z]{1,3}[1-9][0-9]*|R[0-9]+C[0-9]+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            root = "Cupola_" + root;
        root = root.Substring(0, Math.Min(200, root.Length));
        var names = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
        if (!names.Contains(root)) return root;
        for (var index = 2; ; index++)
        {
            var suffix = "_" + index;
            var candidate = root.Substring(0, Math.Min(root.Length, 200 - suffix.Length)) + suffix;
            if (!names.Contains(candidate)) return candidate;
        }
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
