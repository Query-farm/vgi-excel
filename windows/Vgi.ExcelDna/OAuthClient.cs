using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using Newtonsoft.Json;

namespace QueryFarm.Vgi.ExcelDna;

internal static class OAuthClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(2) };
    private static readonly object Gate = new();
    private static readonly Dictionary<string, OAuthTokens> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, bool> ProtectionCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, Task> PendingSignIns = new(StringComparer.OrdinalIgnoreCase);

    static OAuthClient() => ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

    public static bool ShouldPromptForSignIn(Exception error)
    {
        var message = error.ToString();
        if (string.IsNullOrWhiteSpace(message)) return false;
        if (Regex.IsMatch(message, @"token exchange failed|token refresh failed|invalid_grant|AADSTS\d+", RegexOptions.IgnoreCase)) return false;
        return IsAuthenticationFailure(error);
    }

    public static bool IsAuthenticationFailure(Exception error) => Regex.IsMatch(
        error.ToString(),
        @"\b401\b|\b403\b|unauthori[sz]|unauthenticated|authenticat|\boauth\b|invalid[_ ]token|invalid_grant|AADSTS\d+|token[^.]{0,20}(expired|exchange|refresh)|expired[^.]{0,20}token",
        RegexOptions.IgnoreCase);

    public static void PrepareForAttach(VgiConnection connection)
    {
        ConnectionStore.Validate(connection);
        var flowId = "attach-" + Guid.NewGuid().ToString("N");
        var advertised = connection.Authentication == "oauth" || AdvertisesAuthentication(connection, flowId);
        OAuthTraceLog.Write("oauth_attach_prepare", flowId, connection, new { configured_authentication = connection.Authentication, oauth_required = advertised });
        if (!advertised) return;
        connection.Authentication = "oauth";
        ConnectionStore.Save(connection, false);
        EnsureSignedIn(connection);
    }

    public static OAuthAttachCredential GetAttachCredential(VgiConnection connection)
    {
        var tokens = Tokens(connection) ?? throw new InvalidOperationException($"Sign in to the '{connection.Name}' connection from Cupola > Connections.");
        var credential = SelectAttachCredential(tokens);
        OAuthTraceLog.Write("oauth_attach_credential_selected", "attach-" + Guid.NewGuid().ToString("N"), connection,
            new { option = credential.Option, credential_length = credential.Value.Length });
        return credential;
    }

    internal static OAuthAttachCredential SelectAttachCredential(OAuthTokens tokens)
    {
        if (!string.IsNullOrWhiteSpace(tokens.RefreshToken))
            return new OAuthAttachCredential("oauth_refresh_token", tokens.RefreshToken);
        if (!string.IsNullOrWhiteSpace(tokens.Bearer) && tokens.ExpiresAtUtc > DateTime.UtcNow.AddMinutes(1))
            return new OAuthAttachCredential("bearer_token", tokens.Bearer);
        throw new InvalidOperationException("The OAuth session expired. Sign in again.");
    }

    internal static bool ResourceAdvertisesAuthentication(string json)
    {
        var metadata = JsonConvert.DeserializeObject<ResourceMetadata>(json);
        return metadata?.AuthorizationServers is { Length: > 0 };
    }

    internal static string LoopbackRedirect(int port) => $"http://localhost:{port}/oauth-callback.html";

    public static async Task SignInAsync(VgiConnection connection)
    {
        ConnectionStore.Validate(connection);
        var flowId = Guid.NewGuid().ToString("N");
        var started = Stopwatch.StartNew();
        OAuthTraceLog.Write("oauth_sign_in_started", flowId, connection, new { callback_timeout_seconds = 120 });
        try
        {
            var resourceUrl = new Uri(new Uri(connection.Location), "/.well-known/oauth-protected-resource");
            var resource = await GetJson<ResourceMetadata>(resourceUrl);
            OAuthTraceLog.Write("oauth_resource_metadata_loaded", flowId, connection, new
            {
                metadata_origin = resourceUrl.GetLeftPart(UriPartial.Authority), authorization_server_count = resource.AuthorizationServers?.Length ?? 0,
                scope_count = resource.ScopesSupported?.Length ?? 0, has_client_id = !string.IsNullOrWhiteSpace(resource.ClientId)
            });
            var issuer = resource.AuthorizationServers?.FirstOrDefault()
                ?? throw new InvalidOperationException("The VGI service did not advertise an OAuth authorization server.");
            var oidcUrl = new Uri(new Uri(EnsureSlash(issuer)), ".well-known/openid-configuration");
            var oidc = await GetJson<OidcMetadata>(oidcUrl);
            if (string.IsNullOrWhiteSpace(oidc.AuthorizationEndpoint) || string.IsNullOrWhiteSpace(oidc.TokenEndpoint))
                throw new InvalidOperationException("The OAuth server metadata is incomplete.");
            OAuthTraceLog.Write("oauth_authorization_metadata_loaded", flowId, connection, new
            {
                issuer_origin = new Uri(issuer).GetLeftPart(UriPartial.Authority),
                authorization_origin = new Uri(oidc.AuthorizationEndpoint).GetLeftPart(UriPartial.Authority),
                token_origin = new Uri(oidc.TokenEndpoint).GetLeftPart(UriPartial.Authority)
            });

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            try
            {
                var port = ((IPEndPoint)listener.LocalEndpoint).Port;
                var redirect = LoopbackRedirect(port);
                var verifier = RandomBase64Url(64);
                var state = RandomBase64Url(24);
                string challenge;
                using (var sha = SHA256.Create()) challenge = Base64Url(sha.ComputeHash(Encoding.ASCII.GetBytes(verifier)));
                var clientId = string.IsNullOrWhiteSpace(resource.ClientId) ? "vgi-excel" : resource.ClientId;
                var scope = resource.ScopesSupported is { Length: > 0 }
                    ? string.Join(" ", resource.ScopesSupported)
                    : "openid profile offline_access";
                var tokenEndpoint = string.IsNullOrWhiteSpace(resource.TokenEndpoint) ? oidc.TokenEndpoint : resource.TokenEndpoint;
                var authorize = oidc.AuthorizationEndpoint + (oidc.AuthorizationEndpoint.Contains("?") ? "&" : "?") + Form(new Dictionary<string, string>
                {
                    ["response_type"] = "code", ["client_id"] = clientId, ["redirect_uri"] = redirect,
                    ["scope"] = scope, ["state"] = state, ["code_challenge"] = challenge, ["code_challenge_method"] = "S256"
                });
                OAuthTraceLog.Write("oauth_callback_listener_started", flowId, connection, new { port, redirect_path = "/oauth-callback.html" });
                Process.Start(new ProcessStartInfo(authorize) { UseShellExecute = true });
                OAuthTraceLog.Write("oauth_browser_launched", flowId, connection, new { authorization_origin = new Uri(oidc.AuthorizationEndpoint).GetLeftPart(UriPartial.Authority) });
                var accept = listener.AcceptTcpClientAsync();
                if (await Task.WhenAny(accept, Task.Delay(TimeSpan.FromMinutes(2))) != accept)
                    throw new TimeoutException("OAuth sign-in timed out waiting for the browser callback.");
                using var client = await accept;
                var query = await ReadCallbackRequest(client);
                OAuthTraceLog.Write("oauth_callback_received", flowId, connection, new
                {
                    has_code = !string.IsNullOrWhiteSpace(query["code"]), has_state = !string.IsNullOrWhiteSpace(query["state"]),
                    has_error = !string.IsNullOrWhiteSpace(query["error"])
                });
                var display = ResourceDisplay(resource, connection);
                var displayUrl = string.IsNullOrWhiteSpace(resource.Resource) ? connection.Location : resource.Resource;
                try
                {
                    if (!string.Equals(query["state"], state, StringComparison.Ordinal)) throw new InvalidOperationException("OAuth state validation failed.");
                    if (!string.IsNullOrWhiteSpace(query["error"]))
                        throw new InvalidOperationException($"OAuth authorization failed: {query["error"]} — {query["error_description"]}");
                    var code = query["code"] ?? throw new InvalidOperationException("The OAuth callback did not include an authorization code.");
                    OAuthTraceLog.Write("oauth_callback_validated", flowId, connection);
                    var tokens = await Exchange(tokenEndpoint, new Dictionary<string, string>
                    {
                        ["grant_type"] = "authorization_code", ["code"] = code, ["redirect_uri"] = redirect,
                        ["client_id"] = clientId, ["code_verifier"] = verifier
                    }, connection, flowId);
                    tokens.TokenEndpoint = tokenEndpoint;
                    tokens.ClientId = clientId;
                    tokens.Scope = string.IsNullOrWhiteSpace(tokens.Scope) ? scope : tokens.Scope;
                    tokens.UseIdToken = resource.UseIdTokenAsBearer;
                    Save(connection, tokens);
                    var sessionBytes = new FileInfo(OAuthSessionStore.PathFor(Target(connection))).Length;
                    OAuthTraceLog.Write("oauth_session_persisted", flowId, connection, new { encrypted_bytes = sessionBytes, has_refresh_token = !string.IsNullOrWhiteSpace(tokens.RefreshToken) });
                    await WriteCallbackResponse(client, OAuthCallbackPage.Success(display, displayUrl));
                    OAuthTraceLog.Write("oauth_sign_in_completed", flowId, connection, new { elapsed_ms = started.ElapsedMilliseconds });
                }
                catch (Exception error)
                {
                    await TryWriteCallbackError(client, OAuthTraceLog.Redact(error.Message), display, flowId, connection);
                    throw;
                }
            }
            finally { listener.Stop(); }
        }
        catch (Exception error)
        {
            OAuthTraceLog.Write("oauth_sign_in_failed", flowId, connection, new { elapsed_ms = started.ElapsedMilliseconds }, error);
            throw;
        }
    }

    public static string GetBearerToken(VgiConnection connection)
    {
        var flowId = "refresh-" + Guid.NewGuid().ToString("N");
        var tokens = Tokens(connection);
        if (tokens is null) throw new InvalidOperationException($"Sign in to the '{connection.Name}' connection from Cupola > Connections.");
        if (tokens.ExpiresAtUtc > DateTime.UtcNow.AddMinutes(1) && !string.IsNullOrWhiteSpace(tokens.Bearer)) return tokens.Bearer;
        if (string.IsNullOrWhiteSpace(tokens.RefreshToken)) throw new InvalidOperationException("The OAuth session expired. Sign in again.");
        var refreshed = Exchange(tokens.TokenEndpoint, new Dictionary<string, string>
        {
            ["grant_type"] = "refresh_token", ["refresh_token"] = tokens.RefreshToken,
            ["client_id"] = tokens.ClientId, ["scope"] = tokens.Scope
        }, connection, flowId).GetAwaiter().GetResult();
        if (string.IsNullOrWhiteSpace(refreshed.RefreshToken)) refreshed.RefreshToken = tokens.RefreshToken;
        refreshed.TokenEndpoint = tokens.TokenEndpoint;
        refreshed.ClientId = tokens.ClientId;
        refreshed.Scope = string.IsNullOrWhiteSpace(refreshed.Scope) ? tokens.Scope : refreshed.Scope;
        refreshed.UseIdToken = tokens.UseIdToken;
        Save(connection, refreshed);
        return refreshed.Bearer;
    }

    public static bool IsSignedIn(VgiConnection connection)
    {
        lock (Gate) if (Cache.ContainsKey(Key(connection))) return true;
        return LoadSession(connection) is not null;
    }

    public static void SignOut(VgiConnection connection)
    {
        lock (Gate) Cache.Remove(Key(connection));
        OAuthSessionStore.Delete(Target(connection));
        // Remove sessions written by the short-lived Credential Manager implementation.
        CredentialVault.Delete(Target(connection));
        OAuthTraceLog.Write("oauth_signed_out", "signout-" + Guid.NewGuid().ToString("N"), connection);
    }

    private static OAuthTokens? Tokens(VgiConnection connection)
    {
        var key = Key(connection);
        OAuthTokens? tokens;
        lock (Gate) Cache.TryGetValue(key, out tokens);
        return tokens ?? LoadSession(connection);
    }

    private static OAuthTokens? LoadSession(VgiConnection connection)
    {
        var flowId = "session-" + Guid.NewGuid().ToString("N");
        try
        {
            var session = OAuthSessionStore.Read(Target(connection));
            OAuthTraceLog.Write("oauth_session_loaded", flowId, connection, new { found = session is not null, has_refresh_token = !string.IsNullOrWhiteSpace(session?.RefreshToken) });
            return session;
        }
        catch (Exception error)
        {
            OAuthTraceLog.Write("oauth_session_load_failed", flowId, connection, error: error);
            throw;
        }
    }

    private static bool AdvertisesAuthentication(VgiConnection connection, string flowId)
    {
        var key = Key(connection);
        lock (Gate) if (ProtectionCache.TryGetValue(key, out var cached))
        {
            OAuthTraceLog.Write("oauth_protection_cache_hit", flowId, connection, new { advertised = cached });
            return cached;
        }
        bool? advertised = null;
        try
        {
            var url = new Uri(new Uri(connection.Location), "/.well-known/oauth-protected-resource");
            using var response = Http.GetAsync(url).GetAwaiter().GetResult();
            if (response.IsSuccessStatusCode)
                advertised = ResourceAdvertisesAuthentication(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
            else if (response.StatusCode == HttpStatusCode.NotFound)
                advertised = false;
            OAuthTraceLog.Write("oauth_protection_discovery", flowId, connection, new { http_status = (int)response.StatusCode, advertised });
        }
        catch (Exception error)
        {
            OAuthTraceLog.Write("oauth_protection_discovery_failed", flowId, connection, error: error);
        }
        if (advertised.HasValue) lock (Gate) ProtectionCache[key] = advertised.Value;
        return advertised ?? false;
    }

    private static void EnsureSignedIn(VgiConnection connection)
    {
        if (IsSignedIn(connection))
        {
            OAuthTraceLog.Write("oauth_existing_session_reused", "session-" + Guid.NewGuid().ToString("N"), connection);
            return;
        }
        var key = Key(connection);
        Task pending;
        lock (Gate)
        {
            if (!PendingSignIns.TryGetValue(key, out var existing))
            {
                PendingSignIns[key] = pending = SignInAsync(connection);
                OAuthTraceLog.Write("oauth_interactive_sign_in_required", "session-" + Guid.NewGuid().ToString("N"), connection);
            }
            else pending = existing;
        }
        try { pending.GetAwaiter().GetResult(); }
        finally { lock (Gate) if (PendingSignIns.TryGetValue(key, out var current) && ReferenceEquals(current, pending)) PendingSignIns.Remove(key); }
    }

    private static void Save(VgiConnection connection, OAuthTokens tokens)
    {
        if (string.IsNullOrWhiteSpace(tokens.Bearer)) throw new InvalidOperationException("The OAuth response did not contain the required bearer token.");
        lock (Gate) Cache[Key(connection)] = tokens;
        OAuthSessionStore.Write(Target(connection), tokens.PersistentCopy());
    }

    private static async Task<T> GetJson<T>(Uri url)
    {
        using var response = await Http.GetAsync(url);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"OAuth discovery failed ({(int)response.StatusCode}) at {url}.");
        return JsonConvert.DeserializeObject<T>(body) ?? throw new InvalidOperationException("OAuth discovery returned an empty document.");
    }

    private static async Task<OAuthTokens> Exchange(string endpoint, Dictionary<string, string> values, VgiConnection connection, string flowId)
    {
        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps)
            throw new InvalidOperationException("The OAuth token endpoint must use HTTPS.");
        OAuthTraceLog.Write("oauth_token_exchange_started", flowId, connection, new
        {
            token_origin = uri.GetLeftPart(UriPartial.Authority), grant_type = values.TryGetValue("grant_type", out var grant) ? grant : "unknown"
        });
        using var response = await Http.PostAsync(uri, new FormUrlEncodedContent(values));
        var body = await response.Content.ReadAsStringAsync();
        OAuthTraceLog.Write("oauth_token_exchange_response", flowId, connection, new { http_status = (int)response.StatusCode, response_bytes = Encoding.UTF8.GetByteCount(body) });
        if (!response.IsSuccessStatusCode) throw new InvalidOperationException(OAuthFailure((int)response.StatusCode, body));
        var tokens = JsonConvert.DeserializeObject<OAuthTokens>(body) ?? throw new InvalidOperationException("The token endpoint returned an empty response.");
        tokens.ExpiresAtUtc = DateTime.UtcNow.AddSeconds(Math.Max(1, tokens.ExpiresIn));
        OAuthTraceLog.Write("oauth_tokens_received", flowId, connection, new
        {
            has_access_token = !string.IsNullOrWhiteSpace(tokens.AccessToken), access_token_length = tokens.AccessToken?.Length ?? 0,
            has_refresh_token = !string.IsNullOrWhiteSpace(tokens.RefreshToken), refresh_token_length = tokens.RefreshToken?.Length ?? 0,
            has_id_token = !string.IsNullOrWhiteSpace(tokens.IdToken), id_token_length = tokens.IdToken?.Length ?? 0,
            expires_in_seconds = tokens.ExpiresIn
        });
        return tokens;
    }

    private static async Task<System.Collections.Specialized.NameValueCollection> ReadCallbackRequest(TcpClient client)
    {
        var stream = client.GetStream();
        using var reader = new StreamReader(stream, Encoding.ASCII, false, 4096, true);
        var request = await reader.ReadLineAsync() ?? "";
        while (!string.IsNullOrEmpty(await reader.ReadLineAsync())) { }
        var parts = request.Split(' ');
        var target = parts.Length > 1 ? parts[1] : "/";
        return HttpUtility.ParseQueryString(new Uri("http://localhost" + target).Query);
    }

    private static async Task WriteCallbackResponse(TcpClient client, string html)
    {
        var stream = client.GetStream();
        var bytes = Encoding.UTF8.GetBytes(html);
        var headers = Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {bytes.Length}\r\nCache-Control: no-store\r\nContent-Security-Policy: default-src 'none'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'\r\nX-Content-Type-Options: nosniff\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(headers, 0, headers.Length);
        await stream.WriteAsync(bytes, 0, bytes.Length);
    }

    private static async Task TryWriteCallbackError(TcpClient client, string message, string resourceDisplay, string flowId, VgiConnection connection)
    {
        try { await WriteCallbackResponse(client, OAuthCallbackPage.Error(message, resourceDisplay)); }
        catch (Exception responseError) { OAuthTraceLog.Write("oauth_callback_response_failed", flowId, connection, error: responseError); }
    }

    private static string OAuthFailure(int status, string body)
    {
        try
        {
            var value = JsonConvert.DeserializeObject<OAuthError>(body);
            var errorCode = value?.Error;
            if (!string.IsNullOrWhiteSpace(errorCode))
                return OAuthTraceLog.Redact($"OAuth token exchange failed ({status}): {errorCode} — {value?.ErrorDescription}");
        }
        catch { }
        return $"OAuth token exchange failed ({status}).";
    }

    private static string ResourceDisplay(ResourceMetadata resource, VgiConnection connection)
    {
        if (!string.IsNullOrWhiteSpace(resource.ResourceName)) return resource.ResourceName;
        if (Uri.TryCreate(resource.Resource, UriKind.Absolute, out var advertised)) return advertised.Host;
        if (Uri.TryCreate(connection.Location, UriKind.Absolute, out var configured)) return configured.Host;
        return connection.Name;
    }

    private static string Form(Dictionary<string, string> values) => string.Join("&", values.Select(pair => Uri.EscapeDataString(pair.Key) + "=" + Uri.EscapeDataString(pair.Value)));
    private static string EnsureSlash(string value) => value.EndsWith("/") ? value : value + "/";
    private static string RandomBase64Url(int bytes) { var value = new byte[bytes]; using var random = RandomNumberGenerator.Create(); random.GetBytes(value); return Base64Url(value); }
    private static string Base64Url(byte[] value) => Convert.ToBase64String(value).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    private static string Key(VgiConnection connection) => new Uri(connection.Location).GetLeftPart(UriPartial.Path).TrimEnd('/');
    private static string Target(VgiConnection connection) { using var sha = SHA256.Create(); return "QueryFarm/VgiExcel/OAuth/" + BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Key(connection)))).Replace("-", ""); }

    private sealed class ResourceMetadata
    {
        [JsonProperty("resource")] public string Resource { get; set; } = "";
        [JsonProperty("resource_name")] public string ResourceName { get; set; } = "";
        [JsonProperty("authorization_servers")] public string[]? AuthorizationServers { get; set; }
        [JsonProperty("scopes_supported")] public string[]? ScopesSupported { get; set; }
        [JsonProperty("client_id")] public string ClientId { get; set; } = "";
        [JsonProperty("token_endpoint")] public string TokenEndpoint { get; set; } = "";
        [JsonProperty("use_id_token_as_bearer")] public bool UseIdTokenAsBearer { get; set; }
    }
    private sealed class OidcMetadata
    {
        [JsonProperty("authorization_endpoint")] public string AuthorizationEndpoint { get; set; } = "";
        [JsonProperty("token_endpoint")] public string TokenEndpoint { get; set; } = "";
    }
    private sealed class OAuthError
    {
        [JsonProperty("error")] public string Error { get; set; } = "";
        [JsonProperty("error_description")] public string ErrorDescription { get; set; } = "";
    }
}

internal sealed class OAuthAttachCredential
{
    public OAuthAttachCredential(string option, string value) { Option = option; Value = value; }
    public string Option { get; }
    public string Value { get; }
}

internal sealed class OAuthTokens
{
    [JsonProperty("access_token")] public string AccessToken { get; set; } = "";
    [JsonProperty("refresh_token")] public string RefreshToken { get; set; } = "";
    [JsonProperty("id_token")] public string IdToken { get; set; } = "";
    [JsonProperty("expires_in")] public int ExpiresIn { get; set; } = 3600;
    [JsonProperty("scope")] public string Scope { get; set; } = "";
    public string TokenEndpoint { get; set; } = "";
    public string ClientId { get; set; } = "";
    public bool UseIdToken { get; set; }
    public DateTime ExpiresAtUtc { get; set; }
    [JsonIgnore] public string Bearer => UseIdToken ? IdToken : AccessToken;

    public OAuthTokens PersistentCopy() => string.IsNullOrWhiteSpace(RefreshToken) ? this : new OAuthTokens
    {
        RefreshToken = RefreshToken, TokenEndpoint = TokenEndpoint, ClientId = ClientId,
        Scope = Scope, UseIdToken = UseIdToken, ExpiresAtUtc = DateTime.MinValue
    };
}

internal static class OAuthSessionStore
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("QueryFarm.CupolaForExcel.OAuth.v1");
    private static readonly object Gate = new();

    public static void Write(string target, OAuthTokens tokens)
    {
        var json = JsonConvert.SerializeObject(tokens);
        var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(json), Entropy, DataProtectionScope.CurrentUser);
        var path = PathFor(target);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        lock (Gate)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            try
            {
                File.WriteAllBytes(temporary, encrypted);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }

    public static OAuthTokens? Read(string target)
    {
        var path = PathFor(target);
        lock (Gate)
        {
            if (!File.Exists(path)) return null;
            var encrypted = File.ReadAllBytes(path);
            var plain = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
            return JsonConvert.DeserializeObject<OAuthTokens>(Encoding.UTF8.GetString(plain));
        }
    }

    public static void Delete(string target)
    {
        var path = PathFor(target);
        lock (Gate) if (File.Exists(path)) File.Delete(path);
    }

    internal static string PathFor(string target)
    {
        byte[] hash;
        using (var sha = SHA256.Create()) hash = sha.ComputeHash(Encoding.UTF8.GetBytes(target));
        var name = BitConverter.ToString(hash).Replace("-", "") + ".bin";
        return Path.Combine(Path.GetDirectoryName(ConnectionStore.DiagnosticsPath), "oauth-sessions", name);
    }
}

internal static class CredentialVault
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags, Type; public string TargetName; public string? Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize; public IntPtr CredentialBlob; public uint Persist, AttributeCount;
        public IntPtr Attributes; public string? TargetAlias; public string UserName;
    }
    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredWrite(ref NativeCredential credential, uint flags);
    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredDelete(string target, uint type, uint flags);
    [DllImport("advapi32.dll")] private static extern void CredFree(IntPtr buffer);

    public static void Write(string target, OAuthTokens tokens)
    {
        WriteText(target, JsonConvert.SerializeObject(tokens));
    }

    public static void WriteSecret(string target, string value) => WriteText(target, value);

    private static void WriteText(string target, string value)
    {
        var bytes = Encoding.Unicode.GetBytes(value);
        var blob = Marshal.AllocCoTaskMem(bytes.Length);
        try
        {
            Marshal.Copy(bytes, 0, blob, bytes.Length);
            var credential = new NativeCredential { Type = 1, TargetName = target, CredentialBlobSize = (uint)bytes.Length, CredentialBlob = blob, Persist = 2, UserName = Environment.UserName };
            if (!CredWrite(ref credential, 0)) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        finally { Marshal.FreeCoTaskMem(blob); }
    }

    public static OAuthTokens? Read(string target)
    {
        var json = ReadSecret(target);
        return string.IsNullOrWhiteSpace(json) ? null : JsonConvert.DeserializeObject<OAuthTokens>(json!);
    }

    public static string? ReadSecret(string target)
    {
        if (!CredRead(target, 1, 0, out var pointer)) return null;
        try
        {
            var credential = Marshal.PtrToStructure<NativeCredential>(pointer);
            return Marshal.PtrToStringUni(credential.CredentialBlob, (int)credential.CredentialBlobSize / 2);
        }
        finally { CredFree(pointer); }
    }

    public static void Delete(string target) { if (!CredDelete(target, 1, 0) && Marshal.GetLastWin32Error() != 1168) throw new Win32Exception(Marshal.GetLastWin32Error()); }
}
