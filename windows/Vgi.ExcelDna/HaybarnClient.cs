using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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

    public QueryResult QueryResult(string sql, string? connection = null, int? maxRows = null)
    {
        if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required.");
        var definition = ConnectionStore.Resolve(connection);
        ConnectionStore.Validate(definition);
        OAuthClient.PrepareForAttach(definition);
        try { return QueryResultOnce(definition, sql, maxRows); }
        catch (Exception error) when (OAuthClient.ShouldPromptForSignIn(error))
        {
            OAuthTraceLog.Write("oauth_attach_authentication_rejected", "attach-" + Guid.NewGuid().ToString("N"), definition, error: error);
            OAuthClient.SignOut(definition);
            definition.Authentication = "oauth";
            ConnectionStore.Save(definition, false);
            OAuthClient.SignInAsync(definition).GetAwaiter().GetResult();
            return QueryResultOnce(definition, sql, maxRows);
        }
        catch (Exception error) when (definition.Authentication == "oauth" && OAuthClient.IsAuthenticationFailure(error))
        {
            OAuthTraceLog.Write("oauth_attach_authentication_failed_no_retry", "attach-" + Guid.NewGuid().ToString("N"), definition, error: error);
            throw;
        }
    }

    private static QueryResult QueryResultOnce(VgiConnection definition, string sql, int? maxRows)
    {
        var started = Stopwatch.StartNew();
        var output = Execute(BuildScript(definition, AddDescribePrelude(sql)));
        return ParseResult(output, maxRows, started.Elapsed.TotalMilliseconds);
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
        $"Product={ProductInfo.Name} {ProductInfo.Version}; Build={ProductInfo.Build}; Transport=HTTPS only; TimeZone={UserTimeZone.CurrentIanaId()}; Engine={ExecutablePath()}; Extension={ExtensionPath()}; Registry={ConnectionStore.DiagnosticsPath}; OAuthLog={OAuthTraceLog.Path}; AgentLog={AgentTraceLog.Path}; Connections={ConnectionStore.List().Count}";

    internal static string BuildScript(VgiConnection definition, string sql, string? timeZone = null)
    {
        var builder = new StringBuilder();
        var extension = ExtensionPath();
        if (File.Exists(extension)) builder.AppendLine($"LOAD {SqlString(extension)};");
        else { builder.AppendLine("INSTALL vgi FROM community;"); builder.AppendLine("LOAD vgi;"); }
        builder.AppendLine($"SET TimeZone={SqlString(timeZone ?? UserTimeZone.CurrentIanaId())};");
        var options = new List<string> { "TYPE vgi", $"LOCATION {SqlString(definition.Location)}" };
        if (definition.Authentication == "oauth")
        {
            var credential = OAuthClient.GetAttachCredential(definition);
            options.Add($"{credential.Option} {SqlString(credential.Value)}");
        }
        foreach (var option in definition.AttachOptions ?? new Dictionary<string, object?>())
            options.Add($"{option.Key} {SqlLiteral(option.Value)}");
        builder.AppendLine($"ATTACH {SqlString(definition.Catalog)} AS {QuoteIdentifier(definition.Catalog)} ({string.Join(", ", options)});");
        builder.AppendLine(sql);
        return builder.ToString();
    }

    private static string Execute(string script)
    {
        var start = new ProcessStartInfo(ExecutablePath(), "-json")
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Unable to start the bundled Haybarn query process.");
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
        process.StandardInput.Write(script);
        process.StandardInput.Close();
        if (!process.WaitForExit(300_000))
        {
            try { process.Kill(); } catch { }
            throw new TimeoutException("The VGI query exceeded five minutes.");
        }
        var output = outputTask.GetAwaiter().GetResult();
        var error = errorTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0) throw new InvalidOperationException(string.IsNullOrWhiteSpace(error) ? "Haybarn query failed." : Redact(error.Trim()));
        return output;
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
    private static string SqlString(string value) => $"'{value.Replace("'", "''")}'";
    private static string QuoteIdentifier(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    private static string Redact(string value) => Regex.Replace(value,
        "(?i)(bearer_token|oauth_refresh_token)\\s+['\\\"]?[^'\\\"\\s,;]+", "$1 ***");

    private static string ExecutablePath()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_HAYBARN_PATH");
        if (!string.IsNullOrWhiteSpace(configured)) return configured;
        var path = Path.Combine(AddInDirectory(), "haybarn.exe");
        if (!File.Exists(path)) throw new FileNotFoundException("haybarn.exe must be installed beside the VGI XLL.", path);
        return path;
    }

    private static string ExtensionPath()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_EXTENSION_PATH");
        return string.IsNullOrWhiteSpace(configured) ? Path.Combine(AddInDirectory(), "vgi.duckdb_extension") : configured;
    }

    private static string AddInDirectory() => Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? AppDomain.CurrentDomain.BaseDirectory;
}
