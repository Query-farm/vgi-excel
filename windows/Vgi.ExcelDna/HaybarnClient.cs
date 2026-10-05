using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class QueryResult
{
    public QueryColumn[] Columns { get; set; } = Array.Empty<QueryColumn>();
    public object?[][] Rows { get; set; } = Array.Empty<object?[]>();
    public int RowCount { get; set; }
    public bool Truncated { get; set; }
    public double ElapsedMs { get; set; }
}

internal sealed class QueryColumn
{
    public string Name { get; set; } = "";
    public string Type { get; set; } = "";
}

internal sealed class HaybarnClient
{
    public object Query(string sql, string? connection, bool includeHeaders) =>
        Matrix(QueryResult(sql, connection), includeHeaders);

    public object Value(string sql, string? connection)
    {
        var response = QueryResult(sql, connection, 2);
        if (response.Columns.Length != 1 || response.Rows.Length != 1)
            throw new InvalidOperationException("VGI.VALUE requires exactly one row and one column.");
        return response.Rows[0][0] ?? string.Empty;
    }

    public QueryResult QueryResult(string sql, string? connection = null, int? maxRows = null, CancellationToken cancellation = default)
    {
        if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required.");
        var name = ConnectionStore.Resolve(connection).Name;
        return HaybarnSessions.Cache.Query(name, () =>
        {
            var definition = ConnectionStore.Resolve(name);
            var attachments = ConnectionStore.ResolveAttachments(definition);
            foreach (var member in attachments) OAuthClient.PrepareForAttach(member);
            return BuildSessionScript(attachments, "");
        }, sql, maxRows, cancellation);
    }

    internal static string AddDescribePrelude(string sql)
    {
        var query = (sql ?? "").Trim();
        if (query.EndsWith(";", StringComparison.Ordinal)) query = query.Substring(0, query.Length - 1).TrimEnd();
        var masked = Regex.Replace(Regex.Replace(query, @"--[^\r\n]*|/\*[\s\S]*?\*/", " "), "'(?:''|[^'])*'|\"(?:\"\"|[^\"])*\"", " ").Trim();
        if (!Regex.IsMatch(masked, @"^(SELECT|WITH|VALUES|TABLE)\b", RegexOptions.IgnoreCase) || masked.Contains(";")) return sql ?? "";
        return $"DESCRIBE {query};\n{query};";
    }

    public object Call(string functionName, object[] arguments)
    {
        var checkedName = functionName ?? "";
        if (!System.Text.RegularExpressions.Regex.IsMatch(checkedName, @"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*){1,2}$"))
            throw new ArgumentException("The function name must be schema- or catalog-qualified.");
        var matrices = arguments.Select(ToMatrix).ToArray();
        var shape = matrices.FirstOrDefault(matrix => matrix.Length > 1 || matrix[0].Length > 1);
        var rows = shape?.Length ?? 1;
        var columns = shape?[0].Length ?? 1;
        for (var index = 0; index < matrices.Length; index++)
        {
            var matrix = matrices[index];
            if (matrix.Length == 1 && matrix[0].Length == 1)
                matrices[index] = Enumerable.Range(0, rows).Select(_ => Enumerable.Repeat(matrix[0][0], columns).ToArray()).ToArray();
            else if (matrix.Length != rows || matrix.Any(row => row.Length != columns))
                throw new InvalidOperationException("Range arguments to VGI.CALL must have the same shape.");
        }
        var tuples = new List<string>();
        for (var row = 0; row < rows; row++)
        for (var column = 0; column < columns; column++)
            tuples.Add($"({row}, {column}{(matrices.Length == 0 ? "" : ", " + string.Join(", ", matrices.Select(arg => SqlLiteral(arg[row][column]))))})");
        var argNames = Enumerable.Range(0, matrices.Length).Select(index => $"arg_{index}").ToArray();
        var aliases = string.Join(", ", argNames.Select(QuoteIdentifier));
        var inputColumns = string.IsNullOrEmpty(aliases) ? "_row, _column" : $"_row, _column, {aliases}";
        var function = string.Join(".", checkedName.Split('.').Select(QuoteIdentifier));
        var result = QueryResult($"SELECT _row, _column, {function}({aliases}) AS value FROM (VALUES {string.Join(", ", tuples)}) AS input({inputColumns}) ORDER BY 1, 2");
        var matrixResult = new object[rows, columns];
        foreach (var item in result.Rows)
            matrixResult[Convert.ToInt32(item[0], CultureInfo.InvariantCulture), Convert.ToInt32(item[1], CultureInfo.InvariantCulture)] = item[2] ?? "";
        return matrixResult;
    }

    public static string Diagnostics() =>
        $"Product={ProductInfo.Name} {ProductInfo.Version}; Build={ProductInfo.Build}; Transport=HTTPS only; TimeZone={UserTimeZone.CurrentIanaId()}; Engine={NativeLibraryPath()}; SessionMode=Persistent native; Extension={ExtensionPath()}; Registry={ConnectionStore.DiagnosticsPath}; OAuthLog={OAuthTraceLog.Path}; AgentLog={AgentTraceLog.Path}; Connections={ConnectionStore.List().Count}";

    internal static string ProbePrelude()
    {
        var extension = ExtensionPath();
        var load = BundledExtensionLoadScript(extension);
        return load + " SET vgi_http_timeout_seconds=5; SET http_timeout=5; SET http_retries=0;";
    }

    internal static string BuildScript(VgiConnection definition, string sql, string? timeZone = null)
        => BuildSessionScript(new[] { definition }, sql, timeZone);

    internal static string BuildSessionScript(IReadOnlyList<VgiConnection> attachments, string sql, string? timeZone = null, bool probe = false)
    {
        if (attachments.Count == 0) throw new ArgumentException("A session requires at least one catalog.");
        var builder = new StringBuilder();
        var extension = ExtensionPath();
        builder.AppendLine(BundledExtensionLoadScript(extension));
        builder.AppendLine($"SET TimeZone={SqlString(timeZone ?? UserTimeZone.CurrentIanaId())};");
        foreach (var definition in attachments)
        {
            if (probe) builder.AppendLine($"SET vgi_oauth_enabled={(definition.Authentication == "oauth" ? "true" : "false")};");
            var options = new List<string> { "TYPE vgi", $"LOCATION {SqlString(definition.Location)}" };
            if (definition.Authentication == "oauth")
            {
                var credential = OAuthClient.GetAttachCredential(definition);
                options.Add($"{credential.Option} {SqlString(credential.Value)}");
            }
            foreach (var option in definition.AttachOptions ?? new Dictionary<string, object?>())
                options.Add($"{option.Key} {SqlLiteral(option.Value)}");
            builder.AppendLine($"ATTACH {SqlString(definition.Catalog)} AS {QuoteIdentifier(definition.Catalog)} ({string.Join(", ", options)});");
        }
        builder.AppendLine($"USE {QuoteIdentifier(attachments[0].Catalog)};");
        builder.AppendLine(sql);
        return builder.ToString();
    }

    internal static QueryResult ParseResult(string output, int? maxRows, double elapsedMs)
    {
        var arrays = JsonArrays(NormalizeNonFiniteJson(output));
        var data = arrays.LastOrDefault() ?? new JArray();
        var all = data.OfType<JObject>().ToArray();
        var schema = arrays.Take(Math.Max(0, arrays.Count - 1)).LastOrDefault(value => value.OfType<JObject>().Any() && value.OfType<JObject>().All(item => item["column_name"] is not null && item["column_type"] is not null));
        var declared = schema?.OfType<JObject>().Where(item => !string.IsNullOrWhiteSpace(item.Value<string>("column_name"))).ToDictionary(item => item.Value<string>("column_name")!, item => item.Value<string>("column_type") ?? "VARCHAR", StringComparer.Ordinal) ?? new Dictionary<string, string>(StringComparer.Ordinal);
        var names = declared.Count > 0 ? declared.Keys.ToArray() : all.FirstOrDefault()?.Properties().Select(property => property.Name).ToArray() ?? Array.Empty<string>();
        var types = names.Select(name => declared.TryGetValue(name, out var type) ? type : InferType(all, name)).ToArray();
        var decimalNumbers = names.Select((name, index) => IsDecimalType(types[index]) && all.All(row => row[name] is null || row[name]!.Type == JTokenType.Null || IsExcelSafeDecimal(row[name]!.ToString()))).ToArray();
        var limit = maxRows is > 0 ? Math.Min(maxRows.Value, all.Length) : all.Length;
        var rows = all.Take(limit).Select(item => names.Select((name, index) => CellValue(item[name], decimalNumbers[index])).ToArray()).ToArray();
        return new QueryResult
        {
            Columns = names.Select((name, index) => new QueryColumn { Name = name, Type = types[index] }).ToArray(),
            Rows = rows,
            RowCount = all.Length,
            Truncated = limit < all.Length,
            ElapsedMs = elapsedMs
        };
    }

    private static object? CellValue(JToken? value, bool decimalAsNumber = false)
    {
        if (decimalAsNumber && value is not null && value.Type != JTokenType.Null && double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var numeric) && !double.IsNaN(numeric) && !double.IsInfinity(numeric))
            return numeric == 0d ? 0d : numeric;
        return value?.Type switch
    {
        JTokenType.Null or JTokenType.Undefined => null,
        JTokenType.Integer => SafeInteger(value.Value<long>()),
        JTokenType.Float => value.Value<double>(),
        JTokenType.Boolean => value.Value<bool>(),
        JTokenType.String => NormalizeNonFiniteText(value.Value<string>()),
        _ => value?.ToString(Formatting.None)
    };
    }

    private static List<JArray> JsonArrays(string output)
    {
        var matches = Regex.Matches(output, @"(?m)^\s*\[");
        var values = new List<JArray>();
        for (var index = 0; index < matches.Count; index++)
        {
            var start = matches[index].Index;
            var end = index + 1 < matches.Count ? matches[index + 1].Index : output.Length;
            try { values.Add(JArray.Parse(output.Substring(start, end - start).Trim())); } catch { }
        }
        if (values.Count == 0) values.Add(JArray.Parse(output.Trim()));
        return values;
    }

    private static bool IsDecimalType(string type) => Regex.IsMatch(type ?? "", @"^DECIMAL\s*\(", RegexOptions.IgnoreCase);

    internal static bool IsExcelSafeDecimal(string value)
    {
        var match = Regex.Match((value ?? "").Trim(), @"^-?(\d+)(?:\.(\d+))?$");
        if (!match.Success) return false;
        var digits = (match.Groups[1].Value + match.Groups[2].Value).TrimStart('0');
        return Math.Max(1, digits.Length) <= 15 && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && !double.IsNaN(number) && !double.IsInfinity(number);
    }

    private static object SafeInteger(long value) => value >= -9_007_199_254_740_991L && value <= 9_007_199_254_740_991L ? (object)value : value.ToString(CultureInfo.InvariantCulture);
    private static string NormalizeNonFiniteJson(string json) => Regex.Replace(json,
        @"(?i)([:\[,\s])(-?inf(?:inity)?|nan)(?=\s*[,}\]])",
        match => $"{match.Groups[1].Value}\"__cupola_nonfinite_{match.Groups[2].Value.ToLowerInvariant()}__\"");
    private static string? NormalizeNonFiniteText(string? value) => value?.ToLowerInvariant() switch
    {
        "__cupola_nonfinite_nan__" => "NaN",
        "__cupola_nonfinite_inf__" or "__cupola_nonfinite_infinity__" => "Infinity",
        "__cupola_nonfinite_-inf__" or "__cupola_nonfinite_-infinity__" => "-Infinity",
        _ => value
    };
    private static string InferType(JObject[] rows, string name)
    {
        var value = rows.Select(row => row[name]).FirstOrDefault(item => item?.Type != JTokenType.Null);
        if (value?.Type == JTokenType.String)
        {
            var text = value.Value<string>() ?? "";
            if (Regex.IsMatch(text, @"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}(?::?\d{2})?)$")) return "TIMESTAMP WITH TIME ZONE";
            var timestamp = Regex.Match(text, @"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.(\d+))?$");
            if (timestamp.Success) return timestamp.Groups[1].Value.Length > 7 ? "TIMESTAMP_NS" : "TIMESTAMP";
            if (Regex.IsMatch(text, @"^\d{4}-\d{2}-\d{2}$")) return "DATE";
            if (Regex.IsMatch(text, @"^\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}(?::?\d{2})?)$")) return "TIME WITH TIME ZONE";
            var time = Regex.Match(text, @"^\d{2}:\d{2}:\d{2}(?:\.(\d+))?$");
            if (time.Success) return time.Groups[1].Value.Length > 7 ? "TIME_NS" : "TIME";
        }
        return value?.Type switch
        {
            JTokenType.Integer or JTokenType.Float => "NUMBER",
            JTokenType.Boolean => "BOOLEAN",
            JTokenType.Array => "ARRAY",
            JTokenType.Object => "STRUCT",
            _ => "VARCHAR"
        };
    }

    private static object Matrix(QueryResult result, bool includeHeaders)
    {
        var headerRows = includeHeaders ? 1 : 0;
        var matrix = new object[result.Rows.Length + headerRows, result.Columns.Length];
        if (includeHeaders)
            for (var column = 0; column < result.Columns.Length; column++) matrix[0, column] = result.Columns[column].Name;
        for (var row = 0; row < result.Rows.Length; row++)
        for (var column = 0; column < result.Columns.Length; column++) matrix[row + headerRows, column] = result.Rows[row][column] ?? "";
        return matrix;
    }

    private static object?[][] ToMatrix(object value)
    {
        if (value is object[,] range)
            return Enumerable.Range(0, range.GetLength(0)).Select(row => Enumerable.Range(0, range.GetLength(1)).Select(column => Normalize(range[row, column])).ToArray()).ToArray();
        return new[] { new[] { Normalize(value) } };
    }

    private static object? Normalize(object value) => value is ExcelEmpty or ExcelMissing ? null : value;
    private static string SqlLiteral(object? value) => value switch
    {
        null => "NULL",
        bool boolean => boolean ? "TRUE" : "FALSE",
        byte or short or int or long or float or double or decimal => Convert.ToString(value, CultureInfo.InvariantCulture) ?? "NULL",
        _ => SqlString(Convert.ToString(value, CultureInfo.InvariantCulture) ?? "")
    };
    internal static string SqlString(string value) => $"'{value.Replace("'", "''")}'";
    private static string QuoteIdentifier(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    private static string Redact(string value) => Regex.Replace(value,
        "(?i)(bearer_token|oauth_refresh_token)\\s+['\\\"]?[^'\\\"\\s,;]+", "$1 ***");

    internal static string NativeLibraryPath()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_HAYBARN_NATIVE_PATH");
        return Path.GetFullPath(string.IsNullOrWhiteSpace(configured) ? Path.Combine(AddInDirectory(), "haybarn_odbc.dll") : configured);
    }

    private static string ExecutablePath()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_HAYBARN_PATH");
        if (!string.IsNullOrWhiteSpace(configured)) return configured;
        var path = Path.Combine(AddInDirectory(), "haybarn.exe");
        if (!File.Exists(path)) throw new FileNotFoundException("haybarn.exe must be installed beside the VGI XLL.", path);
        return path;
    }

    internal static string BundledExtensionLoadScript(string path)
    {
        if (!File.Exists(path)) throw new InvalidOperationException("A required Cupola component is missing. Repair or reinstall Cupola for Excel.");
        return $"LOAD {SqlString(Path.GetFullPath(path))};";
    }

    private static string ExtensionPath()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_EXTENSION_PATH");
        return string.IsNullOrWhiteSpace(configured) ? Path.Combine(AddInDirectory(), "vgi.duckdb_extension") : configured;
    }

    private static string AddInDirectory() => Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? AppDomain.CurrentDomain.BaseDirectory;
}
