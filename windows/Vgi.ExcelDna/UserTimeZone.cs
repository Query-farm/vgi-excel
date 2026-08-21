using System;
using TimeZoneConverter;

namespace QueryFarm.Vgi.ExcelDna;

internal static class UserTimeZone
{
    internal static string CurrentIanaId() => ToIanaId(TimeZoneInfo.Local.Id);

    internal static string ToIanaId(string timeZoneId)
    {
        var value = (timeZoneId ?? "").Trim();
        if (value.Length == 0) return "Etc/UTC";
        if (value.Contains("/")) return value;
        try { return TZConvert.WindowsToIana(value); }
        catch (TimeZoneNotFoundException) { return "Etc/UTC"; }
        catch (InvalidTimeZoneException) { return "Etc/UTC"; }
    }
}
