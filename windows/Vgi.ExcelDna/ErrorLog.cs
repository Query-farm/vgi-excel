using System;
using System.IO;

namespace QueryFarm.Vgi.ExcelDna;

internal static class ErrorLog
{
    public static string LastMessage { get; private set; } = "No VGI XLL error has been recorded.";

    public static void Write(Exception error)
    {
        var message = OAuthTraceLog.Redact(error.ToString());
        LastMessage = message;
        try
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(local) || local.IndexOf("systemprofile", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                local = string.IsNullOrWhiteSpace(profile) || profile.IndexOf("systemprofile", StringComparison.OrdinalIgnoreCase) >= 0
                    ? Path.Combine(Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\", "Users", Environment.UserName, "AppData", "Local")
                    : Path.Combine(profile, "AppData", "Local");
            }
            var directory = Path.Combine(local, "QueryFarm", "VgiExcel");
            Directory.CreateDirectory(directory);
            File.AppendAllText(
                Path.Combine(directory, "xll.log"),
                $"[{DateTimeOffset.Now:O}] {message}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
            // Diagnostics must never change the worksheet result.
        }
    }
}
