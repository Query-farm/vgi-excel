using System;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryFarm.Vgi.ExcelDna;

internal static class OAuthTraceLog
{
    private static readonly object Gate = new();
    public static string Path => System.IO.Path.Combine(System.IO.Path.GetDirectoryName(ConnectionStore.DiagnosticsPath), "oauth.log");

    public static void Write(string eventName, string flowId, VgiConnection connection, object? details = null, Exception? error = null)
    {
        try
        {
            var payload = details is null ? new JObject() : JObject.FromObject(details);
            payload["timestamp"] = DateTimeOffset.Now.ToString("O");
            payload["event"] = eventName;
            payload["flow_id"] = flowId;
            payload["connection"] = connection.Name;
            payload["catalog"] = connection.Catalog;
            payload["resource_origin"] = new Uri(connection.Location).GetLeftPart(UriPartial.Authority);
            if (error is not null)
            {
                payload["error_type"] = error.GetType().FullName;
                payload["error"] = Redact(error.ToString());
            }
            Sanitize(payload);
            var directory = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            lock (Gate)
            {
                RotateIfNeeded();
                File.AppendAllText(Path, payload.ToString(Formatting.None) + Environment.NewLine);
            }
        }
        catch
        {
            // Authentication must never fail because diagnostics could not be written.
        }
    }

    internal static string Redact(string value)
    {
        var redacted = Regex.Replace(value ?? "", @"(?i)(Authorization:\s*Bearer\s+)[^\s,;]+", "$1***");
        redacted = Regex.Replace(
            redacted,
            @"(?i)([""']?(?:access_token|refresh_token|oauth_refresh_token|id_token|client_secret|authorization_code|code_verifier)[""']?\s*[:=]\s*[""']?)[^""'\s,&;]+",
            "$1***");
        return Regex.Replace(redacted, @"(?i)\b(bearer_token|oauth_refresh_token)\s+[""']?[^""'\s,;]+", "$1 ***");
    }

    private static void Sanitize(JToken token)
    {
        if (token is JObject obj)
        {
            foreach (var property in obj.Properties())
            {
                if (!Regex.IsMatch(property.Name, @"(?i)(^has_|_length$)") &&
                    Regex.IsMatch(property.Name, @"(?i)(access|refresh|id|bearer).*token|token.*(access|refresh|id|bearer)|secret|authorization_code|code_verifier"))
                    property.Value = "***";
                else Sanitize(property.Value);
            }
        }
        else if (token is JArray array)
        {
            for (var index = 0; index < array.Count; index++) Sanitize(array[index]);
        }
        else if (token.Type == JTokenType.String)
        {
            token.Replace(Redact(token.Value<string>() ?? ""));
        }
    }

    private static void RotateIfNeeded()
    {
        if (!File.Exists(Path) || new FileInfo(Path).Length < 2 * 1024 * 1024) return;
        var archive = Path + ".1";
        if (File.Exists(archive)) File.Delete(archive);
        File.Move(Path, archive);
    }
}
