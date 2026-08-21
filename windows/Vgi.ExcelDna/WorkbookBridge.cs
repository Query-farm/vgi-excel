using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using Newtonsoft.Json;

namespace QueryFarm.Vgi.ExcelDna;

internal static class WorkbookBridge
{
    internal const int ExcelWorksheetRows = 1_048_576;
    internal const int ExcelWorksheetColumns = 16_384;
    internal const int MaximumWorksheetDataRows = ExcelWorksheetRows - 1;
    private const int MaxReadCells = 5_000;
    private const string SnapshotPrefix = "_CupolaSnapshot_";
    private static readonly Regex A1Range = new(@"^\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static object Overview()
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        var sheets = new List<object>();
        foreach (dynamic sheet in book.Worksheets)
        {
            dynamic used = sheet.UsedRange;
            var tables = new List<object>();
            foreach (dynamic table in sheet.ListObjects)
                tables.Add(new { name = Convert.ToString(table.Name), address = Address(table.Range), rows = Convert.ToInt32(table.ListRows.Count), columns = Convert.ToInt32(table.ListColumns.Count) });
            var formulaCount = 0;
            try { formulaCount = Convert.ToInt32(used.SpecialCells(-4123).Count); } catch { }
            sheets.Add(new { name = Convert.ToString(sheet.Name), usedRange = Address(used), formulaCount, tables });
        }
        return new { workbook = Convert.ToString(book.Name), activeSheet = Convert.ToString(app.ActiveSheet?.Name), selection = Address(app.Selection), worksheets = sheets };
    }

    public static object ReadRange(string sheetName, string address)
    {
        ValidateA1(address);
        dynamic sheet = Sheet(sheetName);
        dynamic range = sheet.Range[address];
        var rows = Convert.ToInt32(range.Rows.Count);
        var columns = Convert.ToInt32(range.Columns.Count);
        if ((long)rows * columns > MaxReadCells) throw new InvalidOperationException($"The requested range has {(long)rows * columns:N0} cells. Read at most {MaxReadCells:N0} cells at a time.");
        return new { sheet = Convert.ToString(sheet.Name), address = Address(range), values = Matrix(range.Value2, rows, columns), formulas = Matrix(range.Formula, rows, columns) };
    }

    public static object[] ListFormulas(string? sheetName, int requestedLimit)
    {
        var limit = Math.Max(1, Math.Min(500, requestedLimit));
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        var sheets = string.IsNullOrWhiteSpace(sheetName) ? book.Worksheets.Cast<object>().ToArray() : new[] { Sheet(sheetName!) };
        var results = new List<object>();
        foreach (dynamic sheet in sheets)
        {
            dynamic used = sheet.UsedRange;
            var rows = Convert.ToInt32(used.Rows.Count);
            var columns = Convert.ToInt32(used.Columns.Count);
            if ((long)rows * columns > 100_000) throw new InvalidOperationException($"Worksheet “{sheet.Name}” has an unusually large used range. Read a bounded range instead.");
            var formulas = Matrix(used.Formula, rows, columns);
            var values = Matrix(used.Value2, rows, columns);
            for (var row = 0; row < rows && results.Count < limit; row++)
            for (var column = 0; column < columns && results.Count < limit; column++)
            {
                var formula = Convert.ToString(formulas[row][column]) ?? "";
                if (!formula.StartsWith("=", StringComparison.Ordinal)) continue;
                dynamic cell = used.Cells[row + 1, column + 1];
                results.Add(new { sheet = Convert.ToString(sheet.Name), address = Address(cell), formula, value = values[row][column] });
            }
            if (results.Count >= limit) break;
        }
        return results.ToArray();
    }

    public static object InsertAtActiveCell(QueryResult result, string tableName, string? sql = null, string? connection = null)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic sheet = app.ActiveSheet ?? throw new InvalidOperationException("No Excel worksheet is active.");
        dynamic output = WriteTable(result, sheet, app.ActiveCell, UniqueTableName(app.ActiveWorkbook, tableName), null);
        if (!string.IsNullOrWhiteSpace(sql) && !string.IsNullOrWhiteSpace(connection))
            SaveSnapshot(app.ActiveWorkbook, Convert.ToString(output.table) ?? tableName, connection!, sql!);
        return output;
    }

    public static object WriteResult(string mode, QueryResult result, string? sheetName, string tableName)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        if (mode == "new_sheet")
        {
            var validSheetName = UniqueNewSheetName(book, sheetName ?? "VGI Result");
            dynamic sheet = book.Worksheets.Add(After: book.Worksheets[book.Worksheets.Count]);
            sheet.Name = validSheetName;
            return WriteTable(result, sheet, sheet.Range["A1"], UniqueTableName(book, tableName), null);
        }
        if (mode == "replace_table")
        {
            dynamic table = FindTable(book, tableName) ?? throw new InvalidOperationException($"Excel table “{tableName}” was not found.");
            dynamic oldRange = table.Range;
            dynamic sheet = table.Parent;
            dynamic start = oldRange.Cells[1, 1];
            return WriteTable(result, sheet, start, Convert.ToString(table.Name) ?? tableName, table);
        }
        throw new ArgumentException("Workbook write mode must be new_sheet or replace_table.");
    }

    public static bool ActivateTable(string tableName)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        dynamic table = FindTable(book, tableName) ?? throw new InvalidOperationException($"Excel table “{tableName}” was not found.");
        table.Parent.Activate();
        table.Range.Select();
        return true;
    }

    public static object[] ManagedSnapshots()
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        var values = new List<object>();
        foreach (dynamic name in book.Names)
        {
            var localName = Convert.ToString(name.Name)?.Split('!').Last() ?? "";
            if (!localName.StartsWith(SnapshotPrefix, StringComparison.OrdinalIgnoreCase)) continue;
            var metadata = ReadSnapshot(name);
            if (metadata is not null) values.Add(new { metadata.Table, metadata.Connection, metadata.Sql, metadata.UpdatedAt });
        }
        return values.ToArray();
    }

    public static object RefreshSnapshot(string tableName)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        var metadata = FindSnapshot(book, tableName) ?? throw new InvalidOperationException($"Excel table “{tableName}” is not a Cupola-managed snapshot.");
        var result = new HaybarnClient().QueryResult(metadata.Sql, metadata.Connection, MaximumWorksheetDataRows + 1);
        if (result.Truncated || result.RowCount > MaximumWorksheetDataRows)
            throw new InvalidOperationException($"The refreshed result has {result.RowCount:N0} rows. Excel tables can contain at most {MaximumWorksheetDataRows:N0} data rows on a worksheet.");
        var output = WriteResult("replace_table", result, null, tableName);
        SaveSnapshot(book, tableName, metadata.Connection, metadata.Sql);
        return output;
    }

    public static bool ForgetSnapshot(string tableName)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        foreach (dynamic name in book.Names)
        {
            var localName = Convert.ToString(name.Name)?.Split('!').Last() ?? "";
            if (string.Equals(localName, SnapshotName(tableName), StringComparison.OrdinalIgnoreCase)) { name.Delete(); return true; }
        }
        return false;
    }

    public static int RefreshAllSnapshots()
    {
        var snapshots = ManagedSnapshots().Select(value => Convert.ToString(value.GetType().GetProperty("Table")?.GetValue(value))).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
        foreach (var table in snapshots) RefreshSnapshot(table!);
        return snapshots.Length;
    }

    internal static void ValidateA1(string address)
    {
        if (!A1Range.IsMatch((address ?? "").Trim())) throw new ArgumentException("Use a plain A1 range such as A1:F40, without a worksheet or workbook prefix.");
    }

    private static object WriteTable(QueryResult result, dynamic sheet, dynamic start, string tableName, dynamic? existing)
    {
        if (result.Columns.Length == 0) throw new InvalidOperationException("There is no tabular result to insert.");
        ValidateWorksheetBounds(Convert.ToInt32(start.Row), Convert.ToInt32(start.Column), result.Rows.Length + 1, result.Columns.Length);
        var values = Values(result);
        dynamic range = start.Resize[values.GetLength(0), values.GetLength(1)];
        if (existing is not null)
        {
            dynamic oldRange = existing.Range;
            oldRange.ClearContents();
            range.Value2 = values;
            existing.Resize(range);
        }
        else
        {
            range.Value2 = values;
            existing = sheet.ListObjects.Add(1, range, Type.Missing, 1);
            existing.Name = tableName;
        }
        existing.TableStyle = "TableStyleMedium4";
        range.VerticalAlignment = -4160;
        range.EntireColumn.AutoFit();
        for (var column = 1; column <= result.Columns.Length; column++)
        {
            dynamic cells = range.Columns[column];
            if (Convert.ToDouble(cells.ColumnWidth) > 45) cells.ColumnWidth = 45;
            var type = result.Columns[column - 1].Type ?? "";
            if (type.IndexOf("TIMESTAMP", StringComparison.OrdinalIgnoreCase) >= 0) cells.NumberFormat = "yyyy-mm-dd hh:mm:ss.000";
            else if (type.IndexOf("DATE", StringComparison.OrdinalIgnoreCase) >= 0) cells.NumberFormat = "yyyy-mm-dd";
            else if (string.Equals(type, "TIME", StringComparison.OrdinalIgnoreCase)) cells.NumberFormat = "hh:mm:ss.000";
            else if (Regex.IsMatch(type, @"^DECIMAL\s*\(", RegexOptions.IgnoreCase)) cells.NumberFormat = NumberFormat(type, result.Rows, column - 1);
            else if (result.Rows.All(row => row[column - 1] is null || IsInteger(row[column - 1]))) cells.NumberFormat = "#,##0";
            else if (result.Rows.All(row => row[column - 1] is null || row[column - 1] is double or float or decimal)) cells.NumberFormat = "#,##0.########";
        }
        return new { sheet = Convert.ToString(sheet.Name), table = Convert.ToString(existing.Name), address = Address(range) };
    }

    internal static string NumberFormat(string type, object?[][] rows, int column)
    {
        if (rows.Any(row => row[column] is string)) return "@";
        var match = Regex.Match(type ?? "", @"^DECIMAL\s*\(\s*\d+\s*,\s*(\d+)\s*\)", RegexOptions.IgnoreCase);
        if (!match.Success) return "#,##0.########";
        var scale = Math.Min(30, int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture));
        return scale > 0 ? "#,##0." + new string('0', scale) : "#,##0";
    }

    internal static void ValidateWorksheetBounds(int startRow, int startColumn, int rows, int columns)
    {
        if (startRow < 1 || startColumn < 1 || rows < 1 || columns < 1) throw new ArgumentOutOfRangeException("The Excel output range is invalid.");
        if ((long)startRow + rows - 1 > ExcelWorksheetRows)
            throw new InvalidOperationException($"This result does not fit below the selected cell. Excel supports {ExcelWorksheetRows:N0} worksheet rows, including the table header.");
        if ((long)startColumn + columns - 1 > ExcelWorksheetColumns)
            throw new InvalidOperationException($"This result does not fit to the right of the selected cell. Excel supports {ExcelWorksheetColumns:N0} worksheet columns.");
    }

    internal static object[,] Values(QueryResult result)
    {
        var values = new object[result.Rows.Length + 1, result.Columns.Length];
        for (var column = 0; column < result.Columns.Length; column++) values[0, column] = result.Columns[column].Name;
        for (var row = 0; row < result.Rows.Length; row++)
        for (var column = 0; column < result.Columns.Length; column++) values[row + 1, column] = ExcelValue(result.Rows[row][column], result.Columns[column].Type);
        return values;
    }

    internal static object ExcelValue(object? value, string type)
    {
        if (value is null) return "";
        if (string.Equals(type, "TIMESTAMP_NS", StringComparison.OrdinalIgnoreCase) || string.Equals(type, "TIME_NS", StringComparison.OrdinalIgnoreCase) || type.IndexOf("TIME WITH TIME ZONE", StringComparison.OrdinalIgnoreCase) >= 0)
            return value;
        if (value is string text && type.IndexOf("TIMESTAMP WITH TIME ZONE", StringComparison.OrdinalIgnoreCase) >= 0 && DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var offset))
            return DateTime.SpecifyKind(offset.DateTime, DateTimeKind.Unspecified);
        if (value is string plain && (type.IndexOf("DATE", StringComparison.OrdinalIgnoreCase) >= 0 || type.IndexOf("TIMESTAMP", StringComparison.OrdinalIgnoreCase) >= 0) && DateTime.TryParse(plain, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date))
            return date.Year < 1900 ? value : DateTime.SpecifyKind(date, DateTimeKind.Unspecified);
        if (value is string time && string.Equals(type, "TIME", StringComparison.OrdinalIgnoreCase) && TimeSpan.TryParse(time, CultureInfo.InvariantCulture, out var clock))
            return clock.TotalDays;
        return value;
    }

    private static void SaveSnapshot(dynamic book, string table, string connection, string sql)
    {
        ForgetSnapshot(table);
        var metadata = new SnapshotMetadata { Table = table, Connection = connection, Sql = sql, UpdatedAt = DateTime.UtcNow.ToString("O") };
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(metadata)));
        book.Names.Add(Name: SnapshotName(table), RefersTo: $"=\"{encoded}\"", Visible: false);
    }

    private static SnapshotMetadata? FindSnapshot(dynamic book, string table)
    {
        foreach (dynamic name in book.Names)
        {
            var localName = Convert.ToString(name.Name)?.Split('!').Last() ?? "";
            if (string.Equals(localName, SnapshotName(table), StringComparison.OrdinalIgnoreCase)) return ReadSnapshot(name);
        }
        return null;
    }

    private static SnapshotMetadata? ReadSnapshot(dynamic name)
    {
        try
        {
            var refersTo = Convert.ToString(name.RefersTo) ?? "";
            var encoded = refersTo.Trim().TrimStart('=').Trim('"');
            return JsonConvert.DeserializeObject<SnapshotMetadata>(Encoding.UTF8.GetString(Convert.FromBase64String(encoded)));
        }
        catch { return null; }
    }

    private static string SnapshotName(string table) => SnapshotPrefix + Regex.Replace(table, @"[^A-Za-z0-9_]", "_");

    private sealed class SnapshotMetadata
    {
        public string Table { get; set; } = "";
        public string Connection { get; set; } = "";
        public string Sql { get; set; } = "";
        public string UpdatedAt { get; set; } = "";
    }

    private static object?[][] Matrix(object? value, int rows, int columns)
    {
        var result = Enumerable.Range(0, rows).Select(_ => new object?[columns]).ToArray();
        if (value is object[,] matrix)
        {
            for (var row = 0; row < rows; row++) for (var column = 0; column < columns; column++) result[row][column] = matrix[row + 1, column + 1];
        }
        else if (rows > 0 && columns > 0) result[0][0] = value;
        return result;
    }

    private static dynamic Sheet(string name)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic book = app.ActiveWorkbook ?? throw new InvalidOperationException("No Excel workbook is active.");
        foreach (dynamic sheet in book.Worksheets) if (string.Equals(Convert.ToString(sheet.Name), name, StringComparison.OrdinalIgnoreCase)) return sheet;
        throw new InvalidOperationException($"Worksheet “{name}” was not found.");
    }

    private static dynamic? FindTable(dynamic book, string name)
    {
        foreach (dynamic sheet in book.Worksheets) foreach (dynamic table in sheet.ListObjects)
            if (string.Equals(Convert.ToString(table.Name), name, StringComparison.OrdinalIgnoreCase)) return table;
        return null;
    }

    internal static string NormalizeWorksheetName(string? value)
    {
        var name = Regex.Replace((value ?? "").Trim(), @"[\\/\?\*\[\]:\x00-\x1F]", " - ");
        name = Regex.Replace(name, @"\s+", " ").Trim().Trim('\'').Trim();
        if (name.Length == 0) name = "VGI Result";
        if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase)) name = "History Data";
        if (name.Length > 31) name = name.Substring(0, 31).Trim().TrimEnd('\'').Trim();
        return name.Length == 0 ? "VGI Result" : name;
    }

    private static string UniqueNewSheetName(dynamic book, string value)
    {
        var root = NormalizeWorksheetName(value);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (dynamic sheet in book.Worksheets) names.Add(Convert.ToString(sheet.Name) ?? "");
        if (!names.Contains(root)) return root;
        for (var index = 2; ; index++)
        {
            var suffix = $" ({index})";
            var candidate = root.Substring(0, Math.Min(root.Length, 31 - suffix.Length)).TrimEnd() + suffix;
            if (!names.Contains(candidate)) return candidate;
        }
    }

    private static string UniqueTableName(dynamic book, string requested)
    {
        var root = Regex.Replace((requested ?? "VGI_Result").Trim(), @"[^A-Za-z0-9_]", "_");
        if (root.Length == 0 || char.IsDigit(root[0])) root = "VGI_" + root;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (dynamic sheet in book.Worksheets) foreach (dynamic table in sheet.ListObjects) names.Add(Convert.ToString(table.Name) ?? "");
        var value = root; var suffix = 2;
        while (names.Contains(value)) value = root + "_" + suffix++;
        return value;
    }

    private static bool IsInteger(object? value) => value is byte or sbyte or short or ushort or int or uint or long or ulong;
    private static string Address(dynamic value) { try { return Convert.ToString(value?.Address) ?? ""; } catch { return ""; } }
}
