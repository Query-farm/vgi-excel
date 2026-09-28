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
    public string[] Members { get; set; } = Array.Empty<string>();
    [JsonIgnore] public bool IsProfile => Members is { Length: > 0 };
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
            ?? (string.IsNullOrWhiteSpace(name) ? connections.FirstOrDefault() : throw new InvalidOperationException("The requested Cupola connection no longer exists."))
            ?? throw new InvalidOperationException("No VGI connection is configured. Open Cupola > Connections.");
    }

    public static void Save(VgiConnection connection, bool makeDefault = true, string? originalName = null)
    {
        Validate(connection);
        lock (Gate) ConnectionFile.Update(Root, () =>
        {
            if (originalName is not null && !string.Equals(originalName, connection.Name, StringComparison.OrdinalIgnoreCase) && List().Any(item => string.Equals(item.Name, connection.Name, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("A connection with this name already exists. Choose a different name.");
            var values = List().Where(item => !string.Equals(item.Name, connection.Name, StringComparison.OrdinalIgnoreCase)).ToList();
            values.Add(connection);
            foreach (var item in values) ResolveAttachments(item, values);
            ConnectionFile.Write(ConnectionsPath, JsonConvert.SerializeObject(values, Formatting.Indented));
            if (makeDefault) ConnectionFile.Write(DefaultPath, connection.Name);
            return true;
        });
    }

    public static void Remove(string name)
    {
        HaybarnSessions.Cache.Invalidate(name, () =>
        {
            lock (Gate) ConnectionFile.Update(Root, () =>
            {
                var values = List().Where(item => !string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
                if (values.Any(item => (item.Members ?? Array.Empty<string>()).Contains(name, StringComparer.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Remove this connection from its profiles before deleting it.");
                ConnectionFile.Write(ConnectionsPath, JsonConvert.SerializeObject(values, Formatting.Indented));
                if (string.Equals(DefaultName(), name, StringComparison.OrdinalIgnoreCase))
                    ConnectionFile.Write(DefaultPath, values.FirstOrDefault()?.Name ?? "");
                return true;
            });
        });
    }

    public static string? DefaultName()
    {
        lock (Gate) return File.Exists(DefaultPath) ? File.ReadAllText(DefaultPath).Trim() : null;
    }

    public static void SetDefault(string name)
    {
        lock (Gate) ConnectionFile.Update(Root, () =>
        {
            if (!List().Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The selected VGI connection does not exist.");
            ConnectionFile.Write(DefaultPath, name);
            return true;
        });
    }

    internal static int ImportDefaults(string path)
    {
        if (!File.Exists(path)) return 0;
        if (new FileInfo(path).Length > 1024 * 1024) throw new InvalidOperationException("Cupola connection defaults exceed the size limit.");
        var definitions = JsonConvert.DeserializeObject<List<VgiConnection>>(File.ReadAllText(path), new JsonSerializerSettings { MissingMemberHandling = MissingMemberHandling.Error })
            ?? throw new InvalidOperationException("Cupola connection defaults must be an array.");
        foreach (var definition in definitions) Validate(definition);
        if (definitions.Select(item => item.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != definitions.Count)
            throw new InvalidOperationException("Cupola connection defaults contain duplicate names.");
        lock (Gate) return ConnectionFile.Update(Root, () =>
        {
            var values = List().ToList();
            var additions = definitions.Where(item => !values.Any(existing => string.Equals(existing.Name, item.Name, StringComparison.OrdinalIgnoreCase))).ToList();
            if (additions.Count == 0) return 0;
            values.AddRange(additions);
            foreach (var item in values) ResolveAttachments(item, values);
            ConnectionFile.Write(ConnectionsPath, JsonConvert.SerializeObject(values, Formatting.Indented));
            if (string.IsNullOrWhiteSpace(DefaultName())) ConnectionFile.Write(DefaultPath, values[0].Name);
            return additions.Count;
        });
    }

    internal static void ImportMachineDefaults()
    {
        // Tests must never consume a user's or a machine's production configuration.
        if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME"))) return;
        using var machine = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
        using var product = machine.OpenSubKey(@"SOFTWARE\QueryFarm\Cupola");
        var directory = product?.GetValue("InstallDirectory") as string;
        if (!string.IsNullOrWhiteSpace(directory)) ImportDefaults(Path.Combine(directory, "connection-defaults.json"));
    }

    public static void Validate(VgiConnection connection)
    {
        if (string.IsNullOrWhiteSpace(connection.Name)) throw new ArgumentException("A connection name is required.");
        if (connection.IsProfile)
        {
            if (connection.Members.Length > 16 || connection.Members.Any(string.IsNullOrWhiteSpace) || connection.Members.Distinct(StringComparer.OrdinalIgnoreCase).Count() != connection.Members.Length)
                throw new ArgumentException("A profile requires up to 16 distinct saved connections.");
            if (!string.IsNullOrEmpty(connection.Location) || connection.Authentication != "anonymous" || connection.AttachOptions?.Count > 0)
                throw new ArgumentException("Profiles use their members' endpoints, sign-in, and options.");
            return;
        }
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

    internal static IReadOnlyList<VgiConnection> ResolveAttachments(VgiConnection connection, IReadOnlyList<VgiConnection>? saved = null)
    {
        Validate(connection);
        if (!connection.IsProfile) return new[] { connection };
        saved ??= List();
        var attachments = new List<VgiConnection>();
        foreach (var name in connection.Members)
        {
            var matches = saved.Where(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("A profile member is missing or ambiguous. Update the profile in Connections.");
            var member = matches[0];
            if (member.IsProfile) throw new ArgumentException("Profiles can contain saved connections, not other profiles.");
            Validate(member);
            if (attachments.Any(item => string.Equals(item.Catalog, member.Catalog, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("Profile members must have distinct catalog aliases.");
            attachments.Add(member);
        }
        return attachments;
    }

    internal static void InvalidateMemberSessions(string name)
    {
        foreach (var item in List().Where(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase) || (item.Members ?? Array.Empty<string>()).Contains(name, StringComparer.OrdinalIgnoreCase)))
            HaybarnSessions.Cache.Invalidate(item.Name);
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
