using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace QueryFarm.Vgi.ExcelDna;

/// <summary>Direct C API of the pinned Haybarn engine shipped in haybarn_odbc.dll.
/// No ODBC registration, DSN, subprocess, or alternate VGI transport is involved.</summary>
internal sealed class NativeHaybarnSession : IHaybarnSession
{
    private readonly NativeHaybarnApi api = NativeHaybarnApi.Instance.Value;
    private IntPtr database, connection;
    internal NativeHaybarnSession(string setup, CancellationToken cancellation = default)
    {
        cancellation.ThrowIfCancellationRequested();
        try
        {
            if (api.Open(IntPtr.Zero, out database) != 0) throw new InvalidOperationException("Unable to open the native Haybarn engine.");
            if (api.Connect(database, out connection) != 0) throw new InvalidOperationException("Unable to connect to the native Haybarn engine.");
            Query(setup, 1, cancellation);
        }
        catch { Dispose(); throw; }
    }

    public QueryResult Query(string sql, int? maxRows, CancellationToken cancellation = default)
    {
        if (connection == IntPtr.Zero) throw new ObjectDisposedException(nameof(NativeHaybarnSession));
        cancellation.ThrowIfCancellationRequested();
        var started = Stopwatch.StartNew();
        var result = new NativeHaybarnApi.Result();
        var interruptGate = new object();
        var active = true;
        var timedOut = false;
        using var timeout = new Timer(_ =>
        {
            lock (interruptGate) if (active) { timedOut = true; api.Interrupt(connection); }
        }, null, TimeSpan.FromMinutes(5), Timeout.InfiniteTimeSpan);
        using var registration = cancellation.Register(() => { lock (interruptGate) if (active) api.Interrupt(connection); });
        try
        {
            cancellation.ThrowIfCancellationRequested();
            using var text = new NativeHaybarnApi.Utf8(sql);
            var status = api.Query(connection, text.Pointer, ref result);
            lock (interruptGate)
            {
                active = false;
                if (timedOut) throw new TimeoutException("The VGI query exceeded five minutes.");
            }
            if (status != 0 && cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
            if (status != 0) throw new InvalidOperationException(OAuthTraceLog.Redact(NativeHaybarnApi.String(api.Error(ref result))));
            var count = checked((int)api.RowCount(ref result));
            var width = checked((int)api.ColumnCount(ref result));
            var columns = new QueryColumn[width];
            var logicalTypes = new IntPtr[width];
            try
            {
                for (var c = 0; c < width; c++)
                {
                    logicalTypes[c] = api.ColumnLogicalType(ref result, (ulong)c);
                    columns[c] = new QueryColumn { Name = NativeHaybarnApi.String(api.ColumnName(ref result, (ulong)c)), Type = api.TypeName(logicalTypes[c], api.TypeId(logicalTypes[c])) };
                }
                var limit = maxRows is > 0 ? Math.Min(maxRows.Value, count) : count;
                var rows = api.ReadRows(result, limit, logicalTypes, SessionTimeZone());
                return new QueryResult { Columns = columns, Rows = rows, RowCount = count, Truncated = limit < count, ElapsedMs = started.Elapsed.TotalMilliseconds };
            }
            finally { for (var c = 0; c < width; c++) if (logicalTypes[c] != IntPtr.Zero) api.DestroyLogicalType(ref logicalTypes[c]); }

        }
        finally
        {
            // Synchronize with the callback before result/session handles can be freed.
            lock (interruptGate) active = false;
            api.DestroyResult(ref result);
        }
    }

    private TimeZoneInfo SessionTimeZone()
    {
        var result = new NativeHaybarnApi.Result();
        try
        {
            using var sql = new NativeHaybarnApi.Utf8("SELECT current_setting('TimeZone')");
            if (api.Query(connection, sql.Pointer, ref result) != 0) throw new InvalidOperationException("Unable to read the native session time zone.");
            var value = api.ValueString(ref result, 0, 0);
            try { return TimeZoneConverter.TZConvert.GetTimeZoneInfo(NativeHaybarnApi.String(value.Data, checked((int)value.Size))); }
            finally { api.Free(value.Data); }
        }
        finally { api.DestroyResult(ref result); }
    }

    internal static object Cell(string text, int type)
    {
        if (type == 1) return text == "true";
        if ((type >= 2 && type <= 9) || type == 16 || type == 32 || type == 35)
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var integer) && integer >= -9007199254740991L && integer <= 9007199254740991L ? (object)integer : text;
        if (type == 10 || type == 11 || type == 19 && HaybarnClient.IsExcelSafeDecimal(text))
        {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && !double.IsInfinity(number) && !double.IsNaN(number)) return number == 0 ? 0d : number;
            if (text.Equals("nan", StringComparison.OrdinalIgnoreCase)) return "NaN";
            if (text.Equals("inf", StringComparison.OrdinalIgnoreCase)) return "Infinity";
            if (text.Equals("-inf", StringComparison.OrdinalIgnoreCase)) return "-Infinity";
        }
        return text;
    }

    public void Dispose()
    {
        if (connection != IntPtr.Zero) api.Disconnect(ref connection);
        if (database != IntPtr.Zero) api.Close(ref database);
    }
}

internal sealed partial class NativeHaybarnApi
{
    internal static readonly Lazy<NativeHaybarnApi> Instance = new(() => new NativeHaybarnApi());
    // Pin this module for the process lifetime; its code can also be used by Power Query.
    private readonly IntPtr module;
    [StructLayout(LayoutKind.Sequential)] internal struct Result { internal ulong Columns, Rows, Changed; internal IntPtr ColumnData, ErrorMessage, Internal; }
    [StructLayout(LayoutKind.Sequential)] internal struct NativeString { internal IntPtr Data; internal ulong Size; }
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate int OpenDelegate(IntPtr path, out IntPtr database);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate int ConnectDelegate(IntPtr database, out IntPtr connection);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate void DestroyDelegate(ref IntPtr handle);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate void PointerAction(IntPtr handle);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate int QueryDelegate(IntPtr connection, IntPtr sql, ref Result result);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate void DestroyResultDelegate(ref Result result);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate IntPtr ErrorDelegate(ref Result result);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate ulong CountDelegate(ref Result result);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate IntPtr ColumnDelegate(ref Result result, ulong column);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate int TypeDelegate(IntPtr type);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate byte DecimalDelegate(IntPtr type);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] internal delegate NativeString StringDelegate(ref Result result, ulong column, ulong row);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] [return: MarshalAs(UnmanagedType.I1)] internal delegate bool NullDelegate(ref Result result, ulong column, ulong row);
    internal readonly OpenDelegate Open;
    internal readonly ConnectDelegate Connect;
    internal readonly DestroyDelegate Close, Disconnect, DestroyLogicalType;
    internal readonly PointerAction Free, Interrupt;
    internal readonly QueryDelegate Query;
    internal readonly DestroyResultDelegate DestroyResult;
    internal readonly ErrorDelegate Error;
    internal readonly CountDelegate ColumnCount, RowCount;
    internal readonly ColumnDelegate ColumnName, ColumnLogicalType;
    internal readonly TypeDelegate TypeId;
    internal readonly DecimalDelegate DecimalWidth, DecimalScale;
    internal readonly StringDelegate ValueString;
    internal readonly NullDelegate IsNull;
    private NativeHaybarnApi()
    {
        if (IntPtr.Size != 8) throw new PlatformNotSupportedException("Cupola native sessions require 64-bit Excel.");
        var path = HaybarnClient.NativeLibraryPath();
        if (!File.Exists(path)) throw new FileNotFoundException("The bundled Haybarn native engine is missing.", path);
        module = LoadLibraryEx(path, IntPtr.Zero, 0x100 | 0x1000); // DLL directory + safe default directories
        if (module == IntPtr.Zero) throw new InvalidOperationException("Unable to load the bundled Haybarn native engine.");
        InitializeChunks();
        Open = Get<OpenDelegate>("duckdb_open"); Connect = Get<ConnectDelegate>("duckdb_connect");
        Close = Get<DestroyDelegate>("duckdb_close"); Disconnect = Get<DestroyDelegate>("duckdb_disconnect");
        DestroyLogicalType = Get<DestroyDelegate>("duckdb_destroy_logical_type"); Free = Get<PointerAction>("duckdb_free"); Interrupt = Get<PointerAction>("duckdb_interrupt");
        Query = Get<QueryDelegate>("duckdb_query"); DestroyResult = Get<DestroyResultDelegate>("duckdb_destroy_result"); Error = Get<ErrorDelegate>("duckdb_result_error");
        ColumnCount = Get<CountDelegate>("duckdb_column_count"); RowCount = Get<CountDelegate>("duckdb_row_count");
        ColumnName = Get<ColumnDelegate>("duckdb_column_name"); ColumnLogicalType = Get<ColumnDelegate>("duckdb_column_logical_type");
        TypeId = Get<TypeDelegate>("duckdb_get_type_id"); DecimalWidth = Get<DecimalDelegate>("duckdb_decimal_width"); DecimalScale = Get<DecimalDelegate>("duckdb_decimal_scale");
        ValueString = Get<StringDelegate>("duckdb_value_string"); IsNull = Get<NullDelegate>("duckdb_value_is_null");
    }
    private T Get<T>(string name) where T : Delegate
    {
        var address = GetProcAddress(module, name);
        if (address == IntPtr.Zero) throw new InvalidOperationException("The bundled Haybarn native API is incompatible.");
        return Marshal.GetDelegateForFunctionPointer<T>(address);
    }
    internal string TypeName(IntPtr logical, int type) => type == 19 ? $"DECIMAL({DecimalWidth(logical)},{DecimalScale(logical)})" : type switch
    {
        1 => "BOOLEAN", 2 => "TINYINT", 3 => "SMALLINT", 4 => "INTEGER", 5 => "BIGINT", 6 => "UTINYINT", 7 => "USMALLINT", 8 => "UINTEGER", 9 => "UBIGINT", 10 => "FLOAT", 11 => "DOUBLE",
        12 => "TIMESTAMP", 13 => "DATE", 14 => "TIME", 15 => "INTERVAL", 16 => "HUGEINT", 17 => "VARCHAR", 18 => "BLOB", 20 => "TIMESTAMP_S", 21 => "TIMESTAMP_MS", 22 => "TIMESTAMP_NS", 23 => "ENUM", 24 => "LIST", 25 => "STRUCT", 26 => "MAP", 27 => "UUID", 28 => "UNION", 29 => "BIT", 30 => "TIME WITH TIME ZONE", 31 => "TIMESTAMP WITH TIME ZONE", 32 => "UHUGEINT", 33 => "ARRAY", 35 => "BIGNUM", 39 => "TIME_NS", _ => "VARCHAR"
    };
    internal static string String(IntPtr pointer, int? size = null)
    {
        if (pointer == IntPtr.Zero) return "";
        var length = size ?? 0;
        if (!size.HasValue) while (Marshal.ReadByte(pointer, length) != 0) length++;
        var bytes = new byte[length]; Marshal.Copy(pointer, bytes, 0, length); return Encoding.UTF8.GetString(bytes);
    }
    internal sealed class Utf8 : IDisposable
    {
        internal readonly IntPtr Pointer;
        internal Utf8(string text) { var bytes = Encoding.UTF8.GetBytes(text + "\0"); Pointer = Marshal.AllocHGlobal(bytes.Length); Marshal.Copy(bytes, 0, Pointer, bytes.Length); }
        public void Dispose() => Marshal.FreeHGlobal(Pointer);
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)] private static extern IntPtr GetProcAddress(IntPtr module, string name);
}
