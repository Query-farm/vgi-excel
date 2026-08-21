using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class VgiConnection
{
    public string Name { get; set; } = "";
    public string Catalog { get; set; } = "";
    public string Location { get; set; } = "";
    public string Authentication { get; set; } = "anonymous";
    public Dictionary<string, object?> AttachOptions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

internal static class ConnectionStore
{
    private static readonly object Gate = new();
    private static readonly string Root = ConfigRoot();
    // The old localhost companion used connections.json with a different
    // envelope. Keep the direct-XLL registry separate so upgrades are safe.
    private static readonly string ConnectionsPath = Path.Combine(Root, "desktop-connections.json");
    private static readonly string DefaultPath = Path.Combine(Root, "default-connection.txt");

    public static IReadOnlyList<VgiConnection> List()
    {
        lock (Gate)
        {
            if (!File.Exists(ConnectionsPath)) return Array.Empty<VgiConnection>();
            return JsonConvert.DeserializeObject<List<VgiConnection>>(File.ReadAllText(ConnectionsPath))
                ?? new List<VgiConnection>();
        }
    }

    public static VgiConnection Resolve(string? name)
    {
        var connections = List();
        var requested = string.IsNullOrWhiteSpace(name) ? DefaultName() : name;
        return connections.FirstOrDefault(item => string.Equals(item.Name, requested, StringComparison.OrdinalIgnoreCase))
            ?? connections.FirstOrDefault()
            ?? throw new InvalidOperationException("No VGI connection is configured. Open Cupola > Connections.");
    }

    public static void Save(VgiConnection connection, bool makeDefault = true)
    {
        Validate(connection);
        lock (Gate)
        {
            Directory.CreateDirectory(Root);
            var values = List().Where(item => !string.Equals(item.Name, connection.Name, StringComparison.OrdinalIgnoreCase)).ToList();
            values.Add(connection);
            File.WriteAllText(ConnectionsPath, JsonConvert.SerializeObject(values, Formatting.Indented));
            if (makeDefault) File.WriteAllText(DefaultPath, connection.Name);
        }
    }

    public static void Remove(string name)
    {
        lock (Gate)
        {
            Directory.CreateDirectory(Root);
            var values = List().Where(item => !string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
            File.WriteAllText(ConnectionsPath, JsonConvert.SerializeObject(values, Formatting.Indented));
            if (string.Equals(DefaultName(), name, StringComparison.OrdinalIgnoreCase))
                File.WriteAllText(DefaultPath, values.FirstOrDefault()?.Name ?? "");
        }
    }

    public static string? DefaultName()
    {
        lock (Gate) return File.Exists(DefaultPath) ? File.ReadAllText(DefaultPath).Trim() : null;
    }

    public static void SetDefault(string name)
    {
        if (!List().Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("The selected VGI connection does not exist.");
        Directory.CreateDirectory(Root);
        File.WriteAllText(DefaultPath, name);
    }

    public static void Validate(VgiConnection connection)
    {
        if (string.IsNullOrWhiteSpace(connection.Name)) throw new ArgumentException("A connection name is required.");
        if (string.IsNullOrWhiteSpace(connection.Catalog)) throw new ArgumentException("A VGI catalog name is required.");
        if (!Uri.TryCreate(connection.Location, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
            throw new ArgumentException("Cupola for Excel supports HTTPS VGI endpoints only.");
        if (!string.IsNullOrEmpty(uri.UserInfo)) throw new ArgumentException("Credentials must not be embedded in the VGI URL.");
        if (connection.Authentication != "anonymous" && connection.Authentication != "oauth")
            throw new ArgumentException("Authentication must be anonymous or OAuth.");
        var sensitive = new HashSet<string>(new[] { "access_token", "api_key", "authorization", "bearer_token", "client_secret", "id_token", "oauth_refresh_token", "password", "refresh_token", "secret" }, StringComparer.OrdinalIgnoreCase);
        var managed = new HashSet<string>(new[] { "type", "location" }, StringComparer.OrdinalIgnoreCase);
        foreach (var option in connection.AttachOptions ?? new Dictionary<string, object?>())
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(option.Key, @"^[A-Za-z_][A-Za-z0-9_]*$")) throw new ArgumentException($"Invalid ATTACH option name: {option.Key}");
            if (sensitive.Contains(option.Key)) throw new ArgumentException("Credentials must be supplied through VGI sign-in, not ATTACH options.");
            if (managed.Contains(option.Key)) throw new ArgumentException($"{option.Key.ToUpperInvariant()} is managed by Cupola and must not be repeated in ATTACH options.");
            if (option.Value is not null and not string and not bool and not byte and not short and not int and not long and not float and not double and not decimal)
                throw new ArgumentException($"ATTACH option {option.Key} must be a string, number, boolean, or null.");
        }
    }

    public static string DiagnosticsPath => ConnectionsPath;

    private static string LocalAppData()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (!string.IsNullOrWhiteSpace(local) && local.IndexOf("systemprofile", StringComparison.OrdinalIgnoreCase) < 0) return local;
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrWhiteSpace(profile) && profile.IndexOf("systemprofile", StringComparison.OrdinalIgnoreCase) < 0)
            return Path.Combine(profile, "AppData", "Local");
        return Path.Combine(Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\", "Users", Environment.UserName, "AppData", "Local");
    }

    private static string ConfigRoot()
    {
        var testOverride = Environment.GetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME");
        return string.IsNullOrWhiteSpace(testOverride)
            ? Path.Combine(LocalAppData(), "QueryFarm", "VgiExcel")
            : Path.GetFullPath(testOverride);
    }
}
