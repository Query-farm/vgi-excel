using System;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryFarm.Vgi.ExcelDna;

internal static class AgentTraceLog
{
    private const long MaxBytes = 5 * 1024 * 1024;
    private static readonly object Gate = new();
    internal static string Path => System.IO.Path.Combine(Root(), "agent.log");

    public static void Write(JObject value)
    {
        try
        {
            var safe = (JObject)value.DeepClone();
            Redact(safe);
            safe["timestamp"] = DateTimeOffset.Now.ToString("O");
            var line = safe.ToString(Formatting.None) + Environment.NewLine;
            lock (Gate)
            {
                Directory.CreateDirectory(Root());
                if (File.Exists(Path) && new FileInfo(Path).Length >= MaxBytes)
                {
                    var previous = Path + ".1";
                    if (File.Exists(previous)) File.Delete(previous);
                    File.Move(Path, previous);
                }
                File.AppendAllText(Path, line);
            }
        }
        catch
        {
            // Agent diagnostics must never interrupt the user-facing loop.
        }
    }

    private static void Redact(JToken token)
    {
        if (token is JObject value)
        {
            foreach (var property in value.Properties())
            {
                if (Regex.IsMatch(property.Name, "(?i)(api.?key|authorization|credential|secret|token)"))
                    property.Value = "***";
                else Redact(property.Value);
            }
            return;
        }
        if (token is JArray array)
        {
            foreach (var item in array) Redact(item);
            return;
        }
        if (token.Type != JTokenType.String) return;
        var text = token.Value<string>() ?? "";
        text = Regex.Replace(text, "(?i)sk-ant-[A-Za-z0-9_-]+", "sk-ant-***");
        text = Regex.Replace(text, "(?i)((?:api[_-]?key|bearer_token|oauth_refresh_token|refresh_token|authorization|secret)\\s*(?::=|=>|=|:)\\s*['\"]?)[^'\"\\s,;]+", "$1***");
        ((JValue)token).Value = text;
    }

    private static string Root()
    {
        var testOverride = Environment.GetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME");
        if (!string.IsNullOrWhiteSpace(testOverride)) return testOverride;
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return System.IO.Path.Combine(local, "QueryFarm", "VgiExcel");
    }
}
