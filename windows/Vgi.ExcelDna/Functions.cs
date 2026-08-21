using System;
using ExcelDna.Integration;

namespace QueryFarm.Vgi.ExcelDna;

public static class Functions
{
    private static readonly HaybarnClient Client = new();

    [ExcelFunction(Name = "VGI.QUERY", Description = "Runs SQL against a VGI/Haybarn connection and returns a table.")]
    public static object Query(
        [ExcelArgument(Name = "sql")] string sql,
        [ExcelArgument(Name = "connection")] object connection,
        [ExcelArgument(Name = "include_headers")] object includeHeaders,
        [ExcelArgument(Name = "refresh_key")] object refreshKey)
    {
        var name = OptionalString(connection);
        var headers = includeHeaders is ExcelMissing or ExcelEmpty || Convert.ToBoolean(includeHeaders);
        _ = RefreshValue(refreshKey);
        return Safe(() => Client.Query(sql, name, headers));
    }

    [ExcelFunction(Name = "VGI.VALUE", Description = "Runs SQL that returns exactly one value.")]
    public static object Value(
        [ExcelArgument(Name = "sql")] string sql,
        [ExcelArgument(Name = "connection")] object connection,
        [ExcelArgument(Name = "refresh_key")] object refreshKey)
    {
        var name = OptionalString(connection);
        _ = RefreshValue(refreshKey);
        return Safe(() => Client.Value(sql, name));
    }

    [ExcelFunction(Name = "VGI.CALL", Description = "Calls a catalog-qualified VGI scalar function.")]
    public static object Call([ExcelArgument(Name = "function_name")] string functionName, params object[] arguments) =>
        Safe(() => Client.Call(functionName, arguments));

    // Excel 2016-2021 does not understand Office custom-function namespaces.
    // Keep underscore aliases for direct XLL use; the dotted names above are
    // retained for EquivalentAddins conversion in current Windows Excel.
    [ExcelFunction(Name = "VGI_QUERY", Description = "Runs SQL against a VGI/Haybarn connection and returns a table.")]
    public static object QueryLegacy(string sql, object connection, object includeHeaders, object refreshKey) =>
        Query(sql, connection, includeHeaders, refreshKey);

    [ExcelFunction(Name = "VGI_VALUE", Description = "Runs SQL that returns exactly one value.")]
    public static object ValueLegacy(string sql, object connection, object refreshKey) =>
        Value(sql, connection, refreshKey);

    [ExcelFunction(Name = "VGI_CALL", Description = "Calls a catalog-qualified VGI scalar function.")]
    public static object CallLegacy(string functionName, params object[] arguments) => Call(functionName, arguments);

    [ExcelFunction(Name = "VGI_LAST_ERROR", Description = "Returns the latest detailed VGI XLL error for diagnostics.")]
    public static string LastError() => ErrorLog.LastMessage;

    [ExcelFunction(Name = "VGI_DIAGNOSTICS", Description = "Returns the XLL credential path and process identity.")]
    public static string Diagnostics() => HaybarnClient.Diagnostics();

    [ExcelFunction(Name = "VGI_WORKBENCH_STATUS", Description = "Returns the embedded Cupola initialization state.")]
    public static string WorkbenchStatus() => WebWorkbenchForm.LastStatus;

    [ExcelCommand(Name = "VGI_OPEN_WORKBENCH", Description = "Opens Cupola for Excel.")]
    public static void OpenWorkbench() => WorkbenchWindow.Show(2);

    private static string? OptionalString(object value) => value is ExcelMissing or ExcelEmpty ? null : Convert.ToString(value);
    private static object RefreshValue(object value) => value is ExcelMissing or ExcelEmpty ? "" : value;

    private static object Safe(Func<object> action)
    {
        try { return action(); }
        catch (Exception error)
        {
            // Excel-DNA cannot attach a custom message to #N/A. Keep the
            // detailed, redacted error in the XLL log and diagnostics formula.
            ErrorLog.Write(error);
            System.Diagnostics.Trace.TraceError(error.ToString());
            return ExcelError.ExcelErrorNA;
        }
    }
}
