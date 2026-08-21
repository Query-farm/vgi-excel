using System;
using System.Linq;
using System.Text.RegularExpressions;
using Sentry;

namespace QueryFarm.Vgi.ExcelDna;

internal static class SentryTelemetry
{
    internal const string DefaultDsn = "https://f366da21a171a3e200d0a82f2af81e14@o4511299556081664.ingest.us.sentry.io/4511948913049600";
    private const string QueryRedacted = "[query details redacted]";
    private static readonly Regex SqlText = new(@"\b(?:select\s+.+\s+from|with\s+.+\s+as\s*\(|insert\s+into|update\s+.+\s+set|delete\s+from|attach\s+.+\s+as|copy\s+.+\s+(?:to|from)|create\s+(?:or\s+replace\s+)?table|alter\s+table|drop\s+table|pragma\s+|call\s+[\w"".]+\s*\()", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex Credential = new(@"\b(?:Bearer|Basic)\s+[A-Za-z0-9._~+/=-]+|\b(?:sk-ant-|sk-|pk_)[A-Za-z0-9_-]{12,}|\b(?:access[_ -]?token|refresh[_ -]?token|api[_ -]?key|password|secret)\s*[:=]\s*[^\s,;]+", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex Endpoint = new(@"https?://[^\s<>()\[\]{}]+", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex WindowsPath = new(@"\b[A-Za-z]:\\(?:[^\\\r\n]+\\)*[^\\\r\n]*", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex UnixUserPath = new(@"/(?:Users|home)/[^\s:]+", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex QuotedIdentifier = new("[\\\"'“”‘’][^\\\"'“”‘’\\r\\n]{1,160}[\\\"'“”‘’]", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static IDisposable? _sdk;

    internal static bool IsEnabledByConfiguration()
    {
        var setting = Environment.GetEnvironmentVariable("VGI_EXCEL_TELEMETRY");
        return !string.Equals(setting, "0", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(setting, "false", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(setting, "off", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VGI_EXCEL_SENTRY_DSN") ?? DefaultDsn);
    }

    internal static void Initialize()
    {
        if (_sdk is not null || !IsEnabledByConfiguration()) return;
        try
        {
            var dsn = Environment.GetEnvironmentVariable("VGI_EXCEL_SENTRY_DSN") ?? DefaultDsn;
            _sdk = SentrySdk.Init(options =>
            {
                options.Dsn = dsn;
                options.Release = $"cupola-excel@{ProductInfo.Version}+{ProductInfo.Build}";
                options.Distribution = "xll";
                options.Environment = Environment.GetEnvironmentVariable("VGI_EXCEL_SENTRY_ENVIRONMENT") ?? "production";
                options.SendDefaultPii = false;
                options.IsGlobalModeEnabled = true;
                options.AttachStacktrace = true;
                options.TracesSampleRate = 0;
                options.MaxBreadcrumbs = 0;
                options.AutoSessionTracking = false;
                options.EnableLogs = false;
                options.EnableMetrics = false;
                options.DisableFileWrite = true;
                options.ShutdownTimeout = TimeSpan.FromSeconds(2);
                options.DisableAppDomainUnhandledExceptionCapture();
                options.DisableUnobservedTaskExceptionCapture();
                options.DisableAppDomainProcessExitFlush();
                options.DisableNetFxInstallationsIntegration();
                options.DefaultTags["product"] = "cupola-excel";
                options.DefaultTags["host"] = "xll";
                options.DefaultTags["version"] = ProductInfo.Version;
                options.DefaultTags["build"] = ProductInfo.Build;
                options.DefaultTags["transport"] = "https";
                options.SetBeforeBreadcrumb(_ => null!);
                options.SetBeforeSend(ScrubEvent);
            });
        }
        catch
        {
            _sdk = null;
        }
    }

    internal static void Capture(Exception error, string operation)
    {
        if (_sdk is null) return;
        try
        {
            SentrySdk.CaptureException(error, scope => scope.SetTag("operation", SafeOperation(operation)));
        }
        catch
        {
            // Telemetry must never change an Excel operation or its local diagnostics.
        }
    }

    internal static void Shutdown()
    {
        try { _sdk?.Dispose(); }
        catch { }
        finally { _sdk = null; }
    }

    internal static string Redact(string? value)
    {
        var text = OAuthTraceLog.Redact(value ?? "");
        if (text.Length > 4000) text = text.Substring(0, 4000);
        if (SqlText.IsMatch(text)) return QueryRedacted;
        text = Credential.Replace(text, "[credential redacted]");
        text = Endpoint.Replace(text, "[endpoint redacted]");
        text = WindowsPath.Replace(text, "[local path redacted]");
        text = UnixUserPath.Replace(text, "[local path redacted]");
        return QuotedIdentifier.Replace(text, "[identifier redacted]");
    }

    internal static string Classify(string? value)
    {
        var text = value ?? "";
        if (Regex.IsMatch(text, @"abort|cancel|stopp?ed", RegexOptions.IgnoreCase)) return "Operation canceled";
        if (Regex.IsMatch(text, @"timed?\s*out|timeout|did not respond", RegexOptions.IgnoreCase)) return "Operation timed out";
        if (Regex.IsMatch(text, @"\b(?:401|403)\b|oauth|authenticat|unauthori|forbidden|token expired|invalid_grant", RegexOptions.IgnoreCase)) return "Authentication failure";
        if (Regex.IsMatch(text, @"\b(?:429|529)\b|rate.?limit|overload|anthropic", RegexOptions.IgnoreCase)) return "AI provider failure";
        if (Regex.IsMatch(text, @"parser error|syntax error", RegexOptions.IgnoreCase)) return "SQL parser error";
        if (Regex.IsMatch(text, @"binder error|binding error", RegexOptions.IgnoreCase)) return "SQL binder error";
        if (Regex.IsMatch(text, @"catalog|schema|function inventory", RegexOptions.IgnoreCase)) return "Catalog operation failure";
        if (Regex.IsMatch(text, @"worksheet|workbook|excel|range|cell|table name", RegexOptions.IgnoreCase)) return "Excel operation failure";
        if (Regex.IsMatch(text, @"webview|native bridge|xll|excel-dna", RegexOptions.IgnoreCase)) return "Add-in bridge failure";
        if (Regex.IsMatch(text, @"network|fetch|dns|socket|connect|http\s*[45]\d\d", RegexOptions.IgnoreCase)) return "Network failure";
        return "Unexpected application error";
    }

    private static string SafeOperation(string? operation) =>
        operation is not null && Regex.IsMatch(operation, @"^[a-z0-9.-]{1,80}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
            ? operation
            : "unknown";

    private static SentryEvent ScrubEvent(SentryEvent sentryEvent)
    {
        var exceptions = (sentryEvent.SentryExceptions ?? Enumerable.Empty<Sentry.Protocol.SentryException>()).ToArray();
        foreach (var exception in exceptions)
        {
            exception.Value = Classify(exception.Value);
            exception.Mechanism = null;
            if (exception.Stacktrace is null) continue;
            foreach (var frame in exception.Stacktrace.Frames)
            {
                frame.FileName = Redact(frame.FileName);
                frame.AbsolutePath = Redact(frame.AbsolutePath);
                frame.Package = Redact(frame.Package);
                frame.ContextLine = null;
                frame.PreContext.Clear();
                frame.PostContext.Clear();
                frame.Vars.Clear();
            }
        }
        var clean = new SentryEvent
        {
            Level = sentryEvent.Level,
            Logger = "cupola-excel-xll",
            Platform = "csharp",
            Release = $"cupola-excel@{ProductInfo.Version}+{ProductInfo.Build}",
            Distribution = "xll",
            Environment = SafeEnvironment(sentryEvent.Environment),
            SentryExceptions = exceptions,
        };
        clean.SetTag("product", "cupola-excel");
        clean.SetTag("host", "xll");
        clean.SetTag("version", ProductInfo.Version);
        clean.SetTag("build", ProductInfo.Build);
        clean.SetTag("transport", "https");
        if (sentryEvent.Tags.TryGetValue("operation", out var operation)) clean.SetTag("operation", SafeOperation(operation));
        return clean;
    }

    private static string SafeEnvironment(string? environment) =>
        Regex.IsMatch(environment ?? "", @"^[a-z0-9._-]{1,64}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
            ? environment!
            : "production";
}
