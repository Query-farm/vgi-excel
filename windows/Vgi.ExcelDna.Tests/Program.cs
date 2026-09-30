using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using QueryFarm.Vgi.ExcelDna;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var root = Path.Combine(Path.GetTempPath(), "vgi-excel-tests-" + Guid.NewGuid().ToString("N"));
        Environment.SetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME", root);
        Environment.SetEnvironmentVariable("VGI_EXCEL_ANTHROPIC_CREDENTIAL_TARGET", "QueryFarm/VgiExcel/Tests/" + Guid.NewGuid().ToString("N"));
        try
        {
            if (args.Contains("--power-query") || args.Contains("--power-query-profile")) { PowerQueryLiveTests(root, args.Contains("--power-query-profile")); return 0; }
            if (args.Contains("--workbook-tables")) { WorkbookTableTests(root); return 0; }
            if (args.Contains("--native-sessions")) { NativeSessionTests(); return 0; }
            ConnectionProbeTests();
            SessionCacheTests();
            ConnectionPolicyTests(root);
            ConnectionStorageTests(root);
            ProfileTests();
            AgentPolicyTests();
            AgentPromptTests();
            AgentTraceTests(root);
            OAuthTraceTests(root);
            TelemetryPrivacyTests();
            BridgePolicyTests();
            WorkbookPolicyTests();
            PowerQueryTests();
            TimestampRoundTripTests();
            TimeZoneEdgeTests();
            NumericRoundTripTests();
            ExtendedNumericBoundaryTests();
            AccountingDecimalTests();
            RibbonTests();
            WebWorkbenchTests();
            WorkspaceTests();
            Console.WriteLine("PASS: desktop connection and agent policy tests");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine("FAIL: " + error);
            return 1;
        }
        finally
        {
            try { AgentCredentialStore.Delete(); } catch { }
            // WebView2 releases its isolated profile asynchronously after the window closes.
            for (var attempt = 0; Directory.Exists(root); attempt++)
            {
                try { Directory.Delete(root, true); }
                catch (IOException) when (attempt < 100) { System.Threading.Thread.Sleep(100); }
                catch (UnauthorizedAccessException) when (attempt < 100) { System.Threading.Thread.Sleep(100); }
            }
        }
    }

    private static void ConnectionStorageTests(string root)
    {
        var path = Path.Combine(root, "atomic-test.txt");
        ConnectionFile.Write(path, "old");
        ConnectionFile.Write(path, "new");
        Equal("new", File.ReadAllText(path), "atomic connection write replaces the target");
        Equal("old", File.ReadAllText(path + ".bak"), "atomic connection write retains a recovery copy");
        ConnectionFile.Write(path, "0");
        System.Threading.Tasks.Parallel.For(0, 20, _ => ConnectionFile.Update(root, () =>
        {
            ConnectionFile.Write(path, (int.Parse(File.ReadAllText(path)) + 1).ToString());
            return true;
        }));
        Equal("20", File.ReadAllText(path), "independent writers cannot lose connection updates");
        var defaults = Path.Combine(root, "defaults.json");
        var name = "managed-" + Guid.NewGuid().ToString("N");
        var existing = new VgiConnection { Name = "existing-managed-test", Catalog = "existing", Location = "https://existing.test" };
        ConnectionStore.Save(existing);
        File.WriteAllText(defaults, new JArray(
            JObject.FromObject(new VgiConnection { Name = existing.Name.ToUpperInvariant(), Catalog = "different", Location = "https://different.test" }),
            JObject.FromObject(new VgiConnection { Name = name, Catalog = "seed", Location = "https://example.test" })).ToString());
        try
        {
            var selected = ConnectionStore.DefaultName();
            Equal(1, ConnectionStore.ImportDefaults(defaults), "only missing managed defaults are imported");
            Equal(existing.Location, ConnectionStore.Resolve(existing.Name).Location, "IT defaults never overwrite existing user settings");
            Equal(selected, ConnectionStore.DefaultName(), "IT defaults preserve the user's active connection");
            Equal(0, ConnectionStore.ImportDefaults(defaults), "provisioning is idempotent");
            File.WriteAllText(defaults, "[{\"Name\":\"bad\",\"Catalog\":\"bad\",\"Location\":\"https://example.test\",\"Password\":\"secret-fixture\"}]");
            Throws<Newtonsoft.Json.JsonSerializationException>(() => ConnectionStore.ImportDefaults(defaults), "credential properties are rejected in managed configuration");
            File.WriteAllText(defaults, "[{\"Name\":\"bad\",\"Catalog\":\"bad\",\"Location\":\"http://example.test\"}]");
            Throws<ArgumentException>(() => ConnectionStore.ImportDefaults(defaults), "managed configuration requires HTTPS");
            var registry = ConnectionStore.DiagnosticsPath;
            var bytes = File.ReadAllBytes(registry);
            try
            {
                File.WriteAllText(registry, "broken-json");
                Throws<Newtonsoft.Json.JsonReaderException>(() => ConnectionStore.Save(existing), "corrupt registry must never be replaced by an empty store");
                Equal("broken-json", File.ReadAllText(registry), "failed save leaves corrupt registry available for recovery");
            }
            finally { File.WriteAllBytes(registry, bytes); }
        }
        finally { ConnectionStore.Remove(name); ConnectionStore.Remove(existing.Name); }
    }

    private static void ConnectionPolicyTests(string root)
    {
        var weather = new VgiConnection
        {
            Name = "weather", Catalog = "open_meteo",
            Location = "https://vgi-open-meteo.rusty-bb6.workers.dev", Authentication = "anonymous",
            AttachOptions = new System.Collections.Generic.Dictionary<string, object?> { ["region"] = "us-east", ["metadata_cache"] = true }
        };
        ConnectionStore.Save(weather);
        Equal(1, ConnectionStore.List().Count, "saved connection count");
        Equal("weather", ConnectionStore.DefaultName(), "default connection");
        Equal("open_meteo", ConnectionStore.Resolve(null).Catalog, "resolved catalog");
        Equal("us-east", Convert.ToString(ConnectionStore.Resolve(null).AttachOptions["region"]), "attach option persisted");
        var attachScript = HaybarnClient.BuildScript(weather, "SELECT 1;");
        True(attachScript.Contains("region 'us-east'") && attachScript.Contains("metadata_cache TRUE"), "safe attach options included");
        Equal("America/New_York", UserTimeZone.ToIanaId("Eastern Standard Time"), "Windows timezone converted to IANA");
        Equal("America/Los_Angeles", UserTimeZone.ToIanaId("Pacific Standard Time"), "Pacific Windows timezone converted to IANA");
        Equal("America/Denver", UserTimeZone.ToIanaId("America/Denver"), "IANA timezone preserved");
        var zonedScript = HaybarnClient.BuildScript(weather, "SELECT 1;", "America/New_York");
        True(zonedScript.Contains("SET TimeZone='America/New_York';") && zonedScript.IndexOf("SET TimeZone", StringComparison.Ordinal) < zonedScript.IndexOf("ATTACH", StringComparison.Ordinal), "local timezone set before attach");
        True(File.Exists(Path.Combine(root, "desktop-connections.json")), "credential-free registry created");
        var json = File.ReadAllText(Path.Combine(root, "desktop-connections.json"));
        True(!json.Contains("token") && !json.Contains("secret"), "registry contains no credential fields");

        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "http://example.com" }), "HTTP rejected");
        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "uv run worker.py" }), "command rejected");
        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "https://user:secret@example.com" }), "URL credentials rejected");
        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "https://example.com", Authentication = "password" }), "unknown auth rejected");
        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "https://example.com", AttachOptions = new System.Collections.Generic.Dictionary<string, object?> { ["bearer_token"] = "secret" } }), "secret attach option rejected");
        Throws<ArgumentException>(() => ConnectionStore.Validate(new VgiConnection { Name = "bad", Catalog = "bad", Location = "https://example.com", AttachOptions = new System.Collections.Generic.Dictionary<string, object?> { ["bad-key"] = "value" } }), "invalid attach option rejected");
        True(OAuthClient.ShouldPromptForSignIn(new InvalidOperationException("HTTP 401: Authentication required")), "401 requests sign-in");
        True(OAuthClient.ShouldPromptForSignIn(new InvalidOperationException("OAuth token expired")), "expired token requests sign-in");
        True(!OAuthClient.ShouldPromptForSignIn(new InvalidOperationException("Could not reach author.example.com")), "ordinary host error does not request sign-in");
        True(!OAuthClient.ShouldPromptForSignIn(new InvalidOperationException("token exchange failed: invalid_grant")), "IdP rejection is surfaced without a loop");
        True(OAuthClient.IsAuthenticationFailure(new InvalidOperationException("token refresh failed: invalid_grant")), "non-retried token failure is still logged as authentication-related");
        True(!OAuthClient.IsAuthenticationFailure(new InvalidOperationException("Parser Error near SELECT")), "ordinary query failure is excluded from OAuth diagnostics");
        True(OAuthClient.ResourceAdvertisesAuthentication("{\"authorization_servers\":[\"https://login.example.com\"]}"), "protected-resource metadata enables automatic OAuth");
        True(!OAuthClient.ResourceAdvertisesAuthentication("{\"resource\":\"https://public.example.com\"}"), "metadata without an authorization server remains anonymous");
        Equal("http://localhost:54321/oauth-callback.html", OAuthClient.LoopbackRedirect(54321), "Azure-compatible OAuth loopback redirect");
        var attachCredential = OAuthClient.SelectAttachCredential(new OAuthTokens { RefreshToken = "refresh-secret", AccessToken = "access-secret", ExpiresAtUtc = DateTime.UtcNow.AddHours(1) });
        Equal("oauth_refresh_token", attachCredential.Option, "refresh token attach option");
        Equal("refresh-secret", attachCredential.Value, "refresh token attach value");
        var sessionTarget = "QueryFarm/VgiExcel/Tests/OAuth/" + Guid.NewGuid().ToString("N");
        OAuthSessionStore.Write(sessionTarget, new OAuthTokens
        {
            RefreshToken = new string('r', 6000), TokenEndpoint = "https://login.example.com/token",
            ClientId = "cupola-test", Scope = "openid offline_access"
        });
        var restoredSession = OAuthSessionStore.Read(sessionTarget);
        Equal(6000, restoredSession?.RefreshToken.Length ?? 0, "large OAuth session round trip");
        var encryptedSession = File.ReadAllBytes(OAuthSessionStore.PathFor(sessionTarget));
        True(!System.Text.Encoding.UTF8.GetString(encryptedSession).Contains(new string('r', 32)), "OAuth session encrypted at rest");
        OAuthSessionStore.Delete(sessionTarget);
        True(OAuthSessionStore.Read(sessionTarget) is null, "OAuth session deleted");
        ConnectionStore.Remove("weather");
        Equal(0, ConnectionStore.List().Count, "connection removed");
    }

    private static void AgentPolicyTests()
    {
        AgentSqlPolicy.AssertReadOnly("WITH x AS (SELECT 1) SELECT * FROM x;");
        AgentSqlPolicy.AssertReadOnly("SELECT 'DROP TABLE x' AS harmless;");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("DELETE FROM x"), "mutation rejected");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("SELECT 1; DROP TABLE x"), "stacked query rejected");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("WITH x AS (SELECT 1) INSERT INTO t SELECT * FROM x"), "hidden mutation rejected");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("SELECT * FROM read_csv_auto('C:\\Users\\me\\secret.csv')"), "local file reader rejected");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("SELECT * FROM read_parquet('https://example.com/private.parquet')"), "external URL reader rejected");
        Throws<InvalidOperationException>(() => AgentSqlPolicy.AssertReadOnly("SELECT getenv('ANTHROPIC_API_KEY')"), "environment access rejected");
    }

    private static void AgentPromptTests()
    {
        var signedHistory = JArray.Parse("[{'role':'user','content':'hello'},{'role':'assistant','content':[{'type':'thinking','thinking':'','signature':'opaque-signature'},{'type':'redacted_thinking','data':'opaque-data'},{'type':'text','text':'answer'}]}]");
        var request = AgentClient.BuildRequest("claude-sonnet-5", "stable prompt", signedHistory, "max");
        Equal("adaptive", (string?)request["thinking"]?["type"], "native adaptive thinking");
        Equal("max", (string?)request["output_config"]?["effort"], "native model effort");
        True(JToken.DeepEquals(signedHistory, request["messages"]), "native signed thinking is preserved");
        Equal("ephemeral", (string?)request["system"]?[0]?["cache_control"]?["type"], "native system cache prefix");
        True(AgentClient.BuildRequest("claude-haiku-4-5-20251001", "stable", signedHistory)["thinking"] is null, "unsupported thinking is omitted");
        Equal("high", (string?)AgentClient.BuildRequest("claude-opus-5", "stable", signedHistory, "invalid")["output_config"]?["effort"], "invalid effort is normalized");
        foreach (var tableFunction in new[] { "duckdb_tables", "duckdb_views", "duckdb_columns", "duckdb_constraints", "duckdb_schemas", "duckdb_functions", "vgi_function_arguments" })
            AgentSqlPolicy.AssertReadOnly("SELECT * FROM " + tableFunction + "()");

        var connection = new VgiConnection { Name = "weather", Catalog = "open_meteo", Location = "https://secret.example/vgi", Authentication = "oauth" };
        var inventory = new QueryResult
        {
            Columns = new[] { "catalog", "schema", "name", "kind", "description" }.Select(name => new QueryColumn { Name = name, Type = "VARCHAR" }).ToArray(),
            Rows = new[] { new object?[] { "open_meteo", "main", "forecast_current", "table", "Current forecast" } }, RowCount = 1
        };
        var prompt = AgentPromptBuilder.Build(connection, inventory);
        True(prompt.Contains("Connection name: weather"), "agent prompt connection name");
        True(prompt.Contains("Attached catalog: open_meteo"), "agent prompt catalog");
        True(prompt.Contains("`open_meteo.main.forecast_current`"), "agent prompt inventory");
        True(!prompt.Contains(connection.Location), "agent prompt excludes endpoint URL");
    }

    private static void AgentTraceTests(string root)
    {
        AgentTraceLog.Write(new JObject
        {
            ["event"] = "tool_call",
            ["apiKey"] = "top-secret",
            ["input"] = new JObject
            {
                ["sql"] = "SELECT thing(api_key := 'sql-secret')",
                ["authorization"] = "Bearer auth-secret",
                ["oauth_refresh_token"] = "refresh-secret"
            }
        });
        var path = Path.Combine(root, "agent.log");
        True(File.Exists(path), "agent trace log created");
        var text = File.ReadAllText(path);
        True(text.Contains("tool_call") && text.Contains("***"), "agent trace retains useful redacted diagnostics");
        True(!text.Contains("top-secret") && !text.Contains("sql-secret") && !text.Contains("auth-secret") && !text.Contains("refresh-secret"), "agent trace removes credentials");
    }

    private static void OAuthTraceTests(string root)
    {
        var connection = new VgiConnection { Name = "secure", Catalog = "secure", Location = "https://data.example.com/vgi", Authentication = "oauth" };
        OAuthTraceLog.Write("oauth_test_failure", "test-flow", connection, new
        {
            http_status = 400, access_token = "access-secret", access_token_length = 13,
            has_refresh_token = true, authorization_code = "code-secret"
        }, new InvalidOperationException("refresh_token=refresh-secret Authorization: Bearer bearer-secret oauth_refresh_token 'sql-token'"));
        var path = Path.Combine(root, "oauth.log");
        True(File.Exists(path), "OAuth trace log created");
        var trace = File.ReadAllText(path);
        True(trace.Contains("oauth_test_failure") && trace.Contains("test-flow") && trace.Contains("\"http_status\":400"), "OAuth trace retains correlation and status");
        True(trace.Contains("\"access_token_length\":13") && trace.Contains("\"has_refresh_token\":true"), "OAuth trace retains safe token metadata");
        True(!trace.Contains("access-secret") && !trace.Contains("code-secret") && !trace.Contains("refresh-secret") && !trace.Contains("bearer-secret") && !trace.Contains("sql-token"), "OAuth trace removes credentials");

        var discoveryError = OAuthTraceLog.Redact("SELECT * FROM vgi_catalogs('https://example.test', bearer_token := 'bearer''secret value', oauth_refresh_token := 'refresh-secret;tail');");
        True(!discoveryError.Contains("secret") && !discoveryError.Contains("tail") && !discoveryError.Contains("value"), "discovery named SQL credentials are redacted including escaped quotes");
        var success = OAuthCallbackPage.Success("Nearwater <Production>", "https://data.example.com/vgi?a=<unsafe>");
        True(success.Contains("Authentication Successful") && success.Contains("prefers-color-scheme") && success.Contains("Cupola"), "OAuth success page uses the VGI visual treatment");
        True(success.Contains("Nearwater &lt;Production&gt;") && !success.Contains("Nearwater <Production>"), "OAuth success page escapes resource content");
        var failure = OAuthCallbackPage.Error("State mismatch", "Nearwater");
        True(failure.Contains("Authentication Failed") && failure.Contains("icon-circle-error"), "OAuth failure page uses the VGI visual treatment");
    }

    private static void TelemetryPrivacyTests()
    {
        // Exercise the SDK pipeline with an in-memory transport: never contact Sentry.
        var transport = new RecordingTelemetryTransport();
        var options = new Sentry.SentryOptions();
        SentryTelemetry.ConfigureOptions(options);
        options.Dsn = "https://test@example.invalid/1";
        options.Transport = transport;
        True(!options.SendClientReports, "native telemetry disables client reports");
        using (Sentry.SentrySdk.Init(options))
        {
            Sentry.SentrySdk.ConfigureScope(scope =>
            {
                scope.User.Username = "CustomerIdentifier";
                scope.SetExtra("sql", "SELECT salary FROM PayrollPrivate");
                scope.SetTag("connection", "CustomerConnection");
            });
            Sentry.SentrySdk.CaptureException(new InvalidOperationException(
                "SELECT salary FROM PayrollPrivate; Bearer secret-token https://private.example/vgi C:\\Users\\PrivateUser\\Board.xlsx"));
            Sentry.SentrySdk.Logger.LogInfo("PrivateLog secret-token CustomerIdentifier");
            Sentry.SentrySdk.Metrics.EmitCounter("PrivateMetric.CustomerIdentifier", 1);
            Sentry.SentrySdk.FlushAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
        }
        var envelopes = transport.Envelopes.ToArray();
        Equal(1, envelopes.Length, "only the explicit error reaches the transport; logs and metrics are dropped");
        var remote = string.Join("\n", envelopes);
        True(remote.Contains("Unexpected application error"), "remote event uses a fixed error classification");
        foreach (var marker in new[] { "SELECT", "PayrollPrivate", "secret-token", "private.example", "PrivateUser", "Board.xlsx", "CustomerIdentifier", "CustomerConnection", "PrivateLog", "PrivateMetric" })
            True(!remote.Contains(marker), "remote event excludes " + marker);

        var safe = SentryTelemetry.Redact("Bearer abc.def.ghi at https://private.example/vgi from C:\\Users\\person\\Books\\Board.xlsx sheet \"Executive Pay\"");
        True(!safe.Contains("abc.def.ghi") && !safe.Contains("private.example") && !safe.Contains("person") && !safe.Contains("Executive Pay"), "Sentry telemetry removes credentials, endpoints, local paths, and workbook identifiers");
        Equal("[query details redacted]", SentryTelemetry.Redact("Binder error while running SELECT salary FROM payroll WHERE employee = 'Ada'"), "Sentry telemetry removes SQL as a unit");
        Equal("SQL binder error", SentryTelemetry.Classify("Binder Error: column Employee_SSN missing from Payroll"), "Sentry telemetry classifies SQL errors without customer details");
        Equal("Unexpected application error", SentryTelemetry.Classify("Acme North confidential failure marker"), "Sentry telemetry replaces unknown messages with a fixed classification");

        var previous = Environment.GetEnvironmentVariable("VGI_EXCEL_TELEMETRY");
        try
        {
            Environment.SetEnvironmentVariable("VGI_EXCEL_TELEMETRY", "0");
            True(!SentryTelemetry.IsEnabledByConfiguration(), "native telemetry kill switch");
            Environment.SetEnvironmentVariable("VGI_EXCEL_TELEMETRY", "1");
            True(SentryTelemetry.IsEnabledByConfiguration(), "native telemetry enabled by default DSN");
        }
        finally { Environment.SetEnvironmentVariable("VGI_EXCEL_TELEMETRY", previous); }
    }

    private sealed class RecordingTelemetryTransport : Sentry.Extensibility.ITransport
    {
        internal readonly System.Collections.Concurrent.ConcurrentQueue<string> Envelopes = new();

        public async System.Threading.Tasks.Task SendEnvelopeAsync(Sentry.Protocol.Envelopes.Envelope envelope, System.Threading.CancellationToken cancellationToken = default)
        {
            using var stream = new MemoryStream();
            await envelope.SerializeAsync(stream, null, cancellationToken);
            Envelopes.Enqueue(Encoding.UTF8.GetString(stream.ToArray()));
        }
    }

    private static void RibbonTests()
    {
        var xml = new VgiRibbon().GetCustomUI("Microsoft.Excel.Workbook");
        foreach (var expected in new[] { "label='Cupola'", "label='Cupola Data'" })
            True(xml.Contains(expected), "ribbon XML should contain " + expected);
        True(xml.Contains("getImage='GetCupolaImage'"), "primary ribbon control uses the Cupola image callback");
        True(!xml.Contains("imageMso='DatabaseInsert'"), "primary ribbon control no longer uses the generic database icon");
        True(!xml.Contains("RefreshCupolaTables") && !xml.Contains("RefreshFormulas"), "ribbon omits redundant refresh commands");
        True(!xml.Contains("VgiConnections") && !xml.Contains("OpenConnections"), "connection management stays in workbench Settings");
        True(!xml.Contains("VgiDiagnostics") && !xml.Contains("ShowDiagnostics") && !xml.Contains("VgiHelpGroup"), "diagnostics stay in Settings About without an empty ribbon group");
        True(!xml.Contains("Companion"), "ribbon must not expose the retired companion");
    }

    private static void BridgePolicyTests()
    {
        Equal("0.5.0", ProductInfo.Version, "native product version");
        var product = JObject.FromObject(WorkbenchBridge.Invoke("app.info", new JObject()).GetAwaiter().GetResult()!, WorkbenchBridge.Serializer);
        Equal(ProductInfo.Name, product.Value<string>("name"), "bridge product name");
        Equal(ProductInfo.Version, product.Value<string>("version"), "bridge product version");
        Equal(ProductInfo.Build, product.Value<string>("build"), "bridge product build");
        Equal(true, WorkbenchBridge.Invoke("ui.ready", new JObject()).GetAwaiter().GetResult(), "UI ready handshake");
        try
        {
            Equal(null, WorkbenchBridge.Invoke("agent.key.load", new JObject()).GetAwaiter().GetResult(), "missing agent key");
            Equal(true, WorkbenchBridge.Invoke("agent.key.save", new JObject { ["key"] = "test-anthropic-key" }).GetAwaiter().GetResult(), "save agent key");
            Equal("test-anthropic-key", WorkbenchBridge.Invoke("agent.key.load", new JObject()).GetAwaiter().GetResult(), "load agent key");
            Equal(true, WorkbenchBridge.Invoke("agent.key.delete", new JObject()).GetAwaiter().GetResult(), "delete agent key");
            Equal(null, WorkbenchBridge.Invoke("agent.key.load", new JObject()).GetAwaiter().GetResult(), "deleted agent key");
        }
        catch (Win32Exception error) when (error.NativeErrorCode == 1312)
        {
            if (Environment.GetEnvironmentVariable("VGI_EXCEL_REQUIRE_CREDENTIAL_MANAGER") == "1") throw;
            Console.WriteLine("SKIP: Windows Credential Manager is unavailable in this noninteractive logon session.");
        }
        var request = new JObject
        {
            ["sql"] = "SELECT * FROM read_csv_auto('C:\\Users\\me\\secret.csv')",
            ["connection"] = "weather",
            ["agent"] = true
        };
        Throws<InvalidOperationException>(() => WorkbenchBridge.Invoke("query.run", request).GetAwaiter().GetResult(), "native bridge rechecks agent SQL");
        request["agent"] = false; request["queryId"] = Guid.NewGuid().ToString();
        Throws<InvalidOperationException>(() => WorkbenchBridge.Invoke("query.agent", request).GetAwaiter().GetResult(), "cancellable agent queries enforce read-only even with a false agent flag");
    }

    private static void WorkbookPolicyTests()
    {
        WorkbookBridge.ValidateA1("A1:F40");
        WorkbookBridge.ValidateA1("$B$2:$C$9");
        WorkbookBridge.ValidateWorksheetBounds(1, 1, 100_001, 4);
        WorkbookBridge.ValidateWorksheetBounds(1, 1, WorkbookBridge.ExcelWorksheetRows, WorkbookBridge.ExcelWorksheetColumns);
        Throws<ArgumentException>(() => WorkbookBridge.ValidateA1("Sheet1!A1"), "sheet-prefixed range rejected");
        Throws<ArgumentException>(() => WorkbookBridge.ValidateA1("[book.xlsx]Sheet1!A1"), "external range rejected");
        Throws<InvalidOperationException>(() => WorkbookBridge.ValidateWorksheetBounds(2, 1, WorkbookBridge.ExcelWorksheetRows, 1), "worksheet row overflow rejected");
        Throws<InvalidOperationException>(() => WorkbookBridge.ValidateWorksheetBounds(1, 2, 1, WorkbookBridge.ExcelWorksheetColumns), "worksheet column overflow rejected");
        Equal("Revenue - Expense - August - Fi", WorkbookBridge.NormalizeWorksheetName("  Revenue / Expense: August [Final]  "), "invalid and long worksheet name normalized");
        Equal("Forecast", WorkbookBridge.NormalizeWorksheetName("'Forecast'"), "worksheet apostrophes normalized");
        Equal("History Data", WorkbookBridge.NormalizeWorksheetName("History"), "reserved History worksheet name normalized");
        Equal("VGI Result", WorkbookBridge.NormalizeWorksheetName(""), "empty worksheet name uses fallback");
    }

    private static void PowerQueryTests()
    {
        Equal("Forecast_2", PowerQueryBridge.TableName("Forecast", new[] { "forecast" }), "table collisions use an Excel-compatible suffix");
        Equal("Cupola_A1", PowerQueryBridge.TableName("A1", Array.Empty<string>()), "cell references are normalized for table names");
        True(PowerQueryBridge.TableName(new string('x', 250), Array.Empty<string>()).Length <= 200, "AI-proposed table names are bounded");
        Equal("Driver={Cupola for Excel};CupolaConnection={weather};", PowerQueryBridge.ConnectionString("weather", "Cupola for Excel"), "DSN-less Cupola ODBC contract");
        Equal("Driver={Driver}}Name};CupolaConnection={finance}}prod};", PowerQueryBridge.ConnectionString("finance}prod", "Driver}Name"), "ODBC brace escaping");
        var formula = PowerQueryBridge.Formula("SELECT \"amount\" FROM finance.main.ledger WHERE note = 'a'", "finance", "Haybarn VGI");
        True(formula.Contains("Odbc.Query(\"Driver={Haybarn VGI};CupolaConnection={finance};\""), "Power Query uses the configured ODBC driver and Cupola connection identity");
        True(formula.Contains("SELECT \"\"amount\"\" FROM finance.main.ledger"), "Power Query M string escaping");
        True(!formula.Contains("https://") && !formula.Contains("token") && !formula.Contains("secret"), "Power Query formula contains no endpoint or credential material");
        True(!PowerQueryBridge.IsDriverRegistered("Cupola test driver " + Guid.NewGuid().ToString("N")), "missing ODBC driver is detected before Excel creates a workbook query");
        PowerQueryBridge.RequireDriver("Cupola for Excel", new[] { "cupola FOR excel" });
        Throws<InvalidOperationException>(() => PowerQueryBridge.RequireDriver("Cupola for Excel", new[] { "Haybarn Driver", "DuckDB Driver" }), "unrelated drivers do not satisfy the Cupola contract");
        Throws<InvalidOperationException>(() => PowerQueryBridge.ResolveConnection("missing-" + Guid.NewGuid().ToString("N")), "Power Query must not fall back to a different saved connection");
        var existing = new VgiConnection { Name = "power-query-test", Catalog = "test", Location = "https://example.test" };
        ConnectionStore.Save(existing);
        try
        {
            Equal(existing.Name, PowerQueryBridge.ResolveConnection(existing.Name.ToUpperInvariant()).Name, "Power Query resolves the exact saved identity case-insensitively");
            Throws<InvalidOperationException>(() => PowerQueryBridge.ResolveConnection("missing-connection"), "Power Query rejects missing identity even when another connection exists");
        }
        finally { ConnectionStore.Remove(existing.Name); }
        var trickyName = "Query;Data Source=wrong\"";
        var source = new System.Data.OleDb.OleDbConnectionStringBuilder(PowerQueryBridge.MashupSource(trickyName).Substring(6));
        Equal("$Workbook$", source.DataSource, "query names cannot override the Mashup data source");
        Equal(trickyName, (string)source["Location"], "query names are quoted in the Mashup connection string");
        var expected = new VgiConnection { Name = "example", Catalog = "sample", Location = "https://example.test/catalog", AttachOptions = new Dictionary<string, object?> { ["timeout"] = 12, ["mode"] = "test" } };
        var identity = new JObject { ["contract_version"] = 1, ["connection_name"] = expected.Name, ["catalog_alias"] = expected.Catalog, ["location"] = expected.Location, ["authentication"] = expected.Authentication, ["attach_options"] = "{\"mode\":\"test\",\"timeout\":12}" };
        PowerQueryBridge.ValidateIdentity(expected, identity);
        foreach (var field in new[] { "contract_version", "connection_name", "catalog_alias", "location", "authentication", "attach_options" })
        {
            var wrong = (JObject)identity.DeepClone();
            wrong[field] = field == "contract_version" ? new JValue(2) : field == "attach_options" ? new JValue("{}") : new JValue("different");
            Throws<InvalidOperationException>(() => PowerQueryBridge.ValidateIdentity(expected, wrong), "ODBC mismatched " + field + " rejected");
        }
        Throws<InvalidOperationException>(() => PowerQueryBridge.Formula("COPY ledger TO 'C:\\ledger.csv'", "finance"), "Power Query export rejects mutating SQL");
    }

    private static void WorkbookTableTests(string root)
    {
        dynamic? excel = null, book = null;
        var connection = new VgiConnection { Name = "legacy-table-fixture", Catalog = "open_meteo", Location = "https://vgi-open-meteo.rusty-bb6.workers.dev" };
        ConnectionStore.Save(connection);
        try
        {
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
            excel.DisplayAlerts = false;
            book = excel.Workbooks.Add();
            dynamic sheet = book.ActiveSheet;
            var result = new QueryResult { Columns = new[] { new QueryColumn { Name = "value", Type = "BIGINT" } }, Rows = new[] { new object?[] { 7L } }, RowCount = 1 };
            WorkbookBridge.InsertAtActiveCell(result, "StaticFixture", (object)excel);
            Equal(0, WorkbookBridge.ManagedSnapshots((object)book).Length, "new snapshot has no Cupola refresh metadata");
            Equal(0, Convert.ToInt32(book.Connections.Count), "new snapshot has no external connection");
            Equal(7d, Convert.ToDouble(sheet.Range["A2"].Value2), "static snapshot contains query values");
            sheet.Range["D1"].Formula = "=SUM(StaticFixture[value])";
            void SeedLegacy(string sql)
            {
                WorkbookBridge.ForgetSnapshot((object)book, "StaticFixture");
                var metadata = new JObject { ["Table"] = "StaticFixture", ["Connection"] = connection.Name, ["Sql"] = sql, ["UpdatedAt"] = "2026-01-01T00:00:00Z" };
                var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(metadata.ToString()));
                book.Names.Add(Name: "_CupolaSnapshot_StaticFixture", RefersTo: "=\"" + encoded + "\"", Visible: false);
            }
            SeedLegacy("SELECT range AS value, range * 2 AS extra FROM range(3)");
            ConnectionStore.Remove(connection.Name);
            Throws<InvalidOperationException>(() => WorkbookBridge.CreateRefreshableCopy((object)book, "StaticFixture"), "failed migration rejects missing connection");
            Equal(1, WorkbookBridge.ManagedSnapshots((object)book).Length, "failed migration retains legacy metadata");
            Equal(7d, Convert.ToDouble(sheet.Range["A2"].Value2), "failed migration retains original data");
            Equal(0, Convert.ToInt32(book.Queries.Count), "failed migration creates no orphan query");
            ConnectionStore.Save(connection);
            WorkbookBridge.RefreshSnapshot("StaticFixture", (object)excel);
            Equal(3, Convert.ToInt32(sheet.ListObjects["StaticFixture"].ListRows.Count), "legacy refresh grows the table");
            Equal(2, Convert.ToInt32(sheet.ListObjects["StaticFixture"].ListColumns.Count), "legacy refresh grows columns");
            Equal(2d, Convert.ToDouble(sheet.Range["A4"].Value2), "legacy refresh updates values");
            True(Convert.ToString(sheet.Range["D1"].Formula).Contains("StaticFixture"), "legacy refresh preserves formula references");
            var path = Path.Combine(root, "legacy-tables.xlsx");
            book.SaveAs(path, 51); book.Close(false);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)book);
            book = excel.Workbooks.Open(path);
            Equal(1, WorkbookBridge.ManagedSnapshots((object)book).Length, "legacy source survives workbook reopen");
            sheet = book.ActiveSheet;
            SeedLegacy("SELECT 42::BIGINT AS value WHERE false");
            WorkbookBridge.RefreshSnapshot("StaticFixture", (object)excel);
            Equal(0, Convert.ToInt32(sheet.ListObjects["StaticFixture"].ListRows.Count), "legacy refresh supports empty results");
            True(string.IsNullOrEmpty(Convert.ToString(sheet.Range["B1"].Value2)), "removed column header is cleared");
            SeedLegacy("SELECT 42::BIGINT AS value");
            WorkbookBridge.RefreshSnapshot("StaticFixture", (object)excel);
            Equal(1, Convert.ToInt32(sheet.ListObjects["StaticFixture"].ListRows.Count), "legacy refresh shrinks the table");
            Equal(42d, Convert.ToDouble(sheet.Range["A2"].Value2), "legacy refresh after reopening");
            WorkbookBridge.ForgetSnapshot((object)book, "StaticFixture");
            Equal(0, WorkbookBridge.ManagedSnapshots((object)book).Length, "retiring legacy refresh removes only metadata");
            Equal(42d, Convert.ToDouble(sheet.Range["A2"].Value2), "retiring legacy refresh keeps data");
            True(Convert.ToString(sheet.Range["D1"].Formula).Contains("StaticFixture"), "retiring legacy refresh keeps formulas");
            Console.WriteLine("PASS: real Excel static insertion, legacy refresh growth/shrink/reopen, and metadata retirement");
        }
        finally
        {
            try { if ((object?)book != null) book.Close(false); } finally { if ((object?)excel != null) excel.Quit(); }
            if ((object?)book != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)book);
            if ((object?)excel != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)excel);
            HaybarnSessions.Cache.Dispose();
        }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    private static void RecordPowerQueryExcel(object application)
    {
        dynamic excel = application;
        GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out var id);
        if (id == 0) throw new InvalidOperationException("Cannot identify the test-owned Excel process.");
        Console.WriteLine("Test Excel PID: " + id);
        var path = Environment.GetEnvironmentVariable("CUPOLA_TEST_EXCEL_PID_LOG");
        if (!string.IsNullOrWhiteSpace(path))
        {
            using var process = Process.GetProcessById((int)id);
            File.AppendAllText(path, new JObject { ["pid"] = id, ["startedUtc"] = process.StartTime.ToUniversalTime().ToString("O") }.ToString(Newtonsoft.Json.Formatting.None) + Environment.NewLine);
        }
    }

    private static void PowerQueryLiveTests(string testRoot, bool multiCatalog = false)
    {
        var elapsed = Stopwatch.StartNew();
        void Phase(string phase) => Console.WriteLine($"Power Query [{elapsed.Elapsed.TotalSeconds:F3}s]: {phase}");
        var driverPath = Environment.GetEnvironmentVariable("CUPOLA_ODBC_DRIVER_PATH") ?? throw new InvalidOperationException("Set CUPOLA_ODBC_DRIVER_PATH.");
        var name = "Cupola PQ " + Guid.NewGuid().ToString("N");
        var installed = Environment.GetEnvironmentVariable("CUPOLA_TEST_INSTALLED_POWERQUERY") == "1";
        var driverName = installed ? PowerQueryBridge.DefaultDriverName : "Cupola Integration " + Guid.NewGuid().ToString("N");
        var connection = new VgiConnection { Name = name, Catalog = "open_meteo", Location = "https://vgi-open-meteo.rusty-bb6.workers.dev" };
        var fixtures = new List<VgiConnection>();
        if (multiCatalog)
        {
            connection.Name = name + " Weather";
            fixtures.Add(connection);
            var earthquakes = new VgiConnection { Name = name + " Earthquakes", Catalog = "earthquakes", Location = "https://vgi-earthquakes.rusty-bb6.workers.dev" };
            fixtures.Add(earthquakes);
            connection = new VgiConnection { Name = name, Members = new[] { connection.Name, earthquakes.Name } };
        }
        fixtures.Add(connection);
        var fixtureNames = new HashSet<string>(fixtures.Select(item => item.Name), StringComparer.Ordinal);
        var from = multiCatalog ? " FROM earthquakes.main.recent LIMIT 1" : "";
        var actualPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QueryFarm", "VgiExcel", "desktop-connections.json");
        var originalExists = File.Exists(actualPath);
        dynamic? excel = null, book = null;
        using var registry = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
        var registration = @"SOFTWARE\ODBC\ODBCINST.INI\" + driverName;
        try
        {
            if (!installed)
            {
                using (var key = registry.CreateSubKey(registration)) key.SetValue("Driver", Path.GetFullPath(driverPath));
                using (var list = registry.CreateSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers")) list.SetValue(driverName, "Installed");
            }
            Environment.SetEnvironmentVariable(PowerQueryBridge.DriverEnvironmentVariable, driverName);
            foreach (var fixture in fixtures) ConnectionStore.Save(fixture);
            if (originalExists)
            {
                var backups = Path.Combine(Environment.CurrentDirectory, "artifacts", "local-test-backups");
                Directory.CreateDirectory(backups);
                File.Copy(actualPath, Path.Combine(backups, "connections-before-power-query-" + Guid.NewGuid().ToString("N") + ".json"));
            }
            var entries = originalExists ? JArray.Parse(File.ReadAllText(actualPath)) : new JArray();
            foreach (var fixture in fixtures) entries.Add(JObject.FromObject(fixture));
            Directory.CreateDirectory(Path.GetDirectoryName(actualPath)!);
            File.WriteAllText(actualPath, entries.ToString());
            Phase("saved fixtures ready; validating direct ODBC identity");
            PowerQueryBridge.ValidateOdbcConnection(connection, driverName);
            Phase("direct ODBC identity passed");
            // A registered but incompatible standard Haybarn driver must fail before any workbook write.
            if (PowerQueryBridge.IsDriverRegistered("Haybarn Driver"))
                Throws<InvalidOperationException>(() => PowerQueryBridge.ValidateOdbcConnection(connection, "Haybarn Driver"), "generic driver cannot validate Cupola identity");
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
            RecordPowerQueryExcel((object)excel);
            var interactive = Environment.GetEnvironmentVariable("CUPOLA_TEST_INTERACTIVE_POWERQUERY") == "1";
            excel.Visible = interactive;
            excel.DisplayAlerts = interactive;
            book = excel.Workbooks.Add();
            const string requestedSheet = "AI / Forecast: Results [Refresh]";
            var normalizedSheet = WorkbookBridge.NormalizeWorksheetName(requestedSheet);
            dynamic existingSheet = book.Worksheets[1];
            existingSheet.Name = normalizedSheet;
            existingSheet.Range["A1"].Value2 = "Existing";
            existingSheet.Range["A2"].Value2 = "Keep me";
            dynamic existingTable = existingSheet.ListObjects.Add(1, existingSheet.Range["A1:A2"], Type.Missing, 1);
            existingTable.Name = "Cupola_AI_Output";
            Phase("test workbook opened; creating query and starting initial refresh");
            var sql = "SELECT open_meteo.main.weather_code_text(0) AS condition" + from;
            var outcome = JObject.FromObject(PowerQueryBridge.CreateInWorkbook(book, sql, name, "Cupola Refresh Test", true, requestedSheet, "Cupola_AI_Output"));
            Phase("query creation returned; awaiting initial worksheet result");
            if (interactive) Console.WriteLine("If Excel requests authentication, choose Default or Custom and leave credentials blank.");
            True(outcome.Value<bool>("loaded"), "Power Query initial load: " + outcome.Value<string>("message"));
            Equal(normalizedSheet.Substring(0, Math.Min(27, normalizedSheet.Length)) + " (2)", outcome.Value<string>("sheet"), "AI sheet name is normalized and made unique");
            Equal("Cupola_AI_Output_2", outcome.Value<string>("table"), "AI table name is made unique without replacing the original");
            Equal("Keep me", (string)existingSheet.Range["A2"].Value2, "existing table is preserved");
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)existingTable);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)existingSheet);
            dynamic sheet = book.Worksheets[(string)outcome["sheet"]!];
            for (var attempt = 0; attempt < (interactive ? 6000 : 1200) && Convert.ToString(sheet.Range["A2"].Value2) != "Clear sky"; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(100); }
            Equal("Clear sky", (string)sheet.Range["A2"].Value2, "Power Query initial VGI result");
            Phase("initial worksheet result passed; changing query for Refresh All");
            dynamic query = book.Queries[(string)outcome["query"]!];
            query.Formula = PowerQueryBridge.Formula("SELECT open_meteo.main.weather_code_text(61) AS condition" + from, name, driverName);
            book.RefreshAll();
            excel.CalculateUntilAsyncQueriesDone();
            Equal("Slight rain", (string)sheet.Range["A2"].Value2, "Excel Refresh All reruns the VGI query");
            Phase("Refresh All result passed; saving workbook and restarting Excel");
            var workbookPath = Path.Combine(testRoot, "power-query-reopen.xlsx");
            excel.DisplayAlerts = false;
            book.SaveAs(workbookPath, 51);
            book.Close(false);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)book);
            book = null;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)query);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)sheet);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)excel);
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
            RecordPowerQueryExcel((object)excel);
            excel.Visible = interactive;
            excel.DisplayAlerts = interactive;
            book = excel.Workbooks.Open(workbookPath);
            Phase("workbook reopened; changing query for another Refresh All");
            dynamic reopenedQuery = book.Queries[(string)outcome["query"]!];
            reopenedQuery.Formula = PowerQueryBridge.Formula("SELECT open_meteo.main.weather_code_text(0) AS condition" + from, name, driverName);
            book.RefreshAll();
            excel.CalculateUntilAsyncQueriesDone();
            dynamic reopenedSheet = book.Worksheets[(string)outcome["sheet"]!];
            Equal("Clear sky", (string)reopenedSheet.Range["A2"].Value2, "new Excel session reattaches the VGI catalog");
            Phase("reopened workbook refresh passed");
            Console.WriteLine("PASS: real Excel Power Query load, Refresh All, and workbook reopen in a new Excel session through Cupola ODBC");
        }
        catch (Exception error) { Console.Error.WriteLine("Power Query test failed: " + error.Message); throw; }
        finally
        {
            Phase("cleanup starting");
            try { if ((object?)excel != null) excel.DisplayAlerts = false; } catch (System.Runtime.InteropServices.COMException) { }
            try { if ((object?)book != null) { foreach (dynamic worksheet in book.Worksheets) foreach (dynamic table in worksheet.QueryTables) table.CancelRefresh(); } } catch (System.Runtime.InteropServices.COMException) { }
            try { if ((object?)book != null) book.Close(false); } catch (System.Runtime.InteropServices.COMException) { }
            try { if ((object?)excel != null) excel.Quit(); } catch (System.Runtime.InteropServices.COMException) { }
            if ((object?)book != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)book);
            if ((object?)excel != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)excel);
            if (!installed)
            {
                registry.DeleteSubKeyTree(registration, false);
                using (var list = registry.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", true)) list?.DeleteValue(driverName, false);
            }
            if (File.Exists(actualPath))
            {
                var entries = JArray.Parse(File.ReadAllText(actualPath));
                foreach (var item in entries.Where(item => fixtureNames.Contains(item.Value<string>("Name") ?? "")).ToArray()) item.Remove();
                if (!originalExists && entries.Count == 0) File.Delete(actualPath); else File.WriteAllText(actualPath, entries.ToString());
            }
            Environment.SetEnvironmentVariable(PowerQueryBridge.DriverEnvironmentVariable, null);
            Phase("cleanup finished");
        }
    }

    private static void TimestampRoundTripTests()
    {
        const string fallback = "[{\"leap_date\":\"2024-02-29\",\"clock\":\"23:59:59.123456\",\"clock_tz\":\"23:59:59.123456-07\",\"timestamp_s\":\"2026-08-19 22:15:30\",\"timestamp_ms\":\"2026-08-19 22:15:30.123\",\"timestamp_us\":\"2026-08-19 22:15:30.123456\",\"timestamp_ns\":\"2026-08-19 22:15:30.123456789\",\"summer\":\"2026-08-19 18:15:30.123456-04\",\"winter\":\"2026-01-19 17:15:30.123456-05\",\"epoch_value\":\"1970-01-01 00:00:00\",\"positive_infinity\":\"infinity\",\"negative_infinity\":\"-infinity\",\"duration\":\"1 year 2 months 3 days 04:05:06.123456\",\"before_excel\":\"1899-12-31\",\"excel_max\":\"9999-12-31\"}]";
        var output = RunHaybarnOrFallback("SET TimeZone='America/New_York';\nSELECT DATE '2024-02-29' AS leap_date, TIME '23:59:59.123456' AS clock, TIMETZ '23:59:59.123456-07:00' AS clock_tz, TIMESTAMP_S '2026-08-19 22:15:30' AS timestamp_s, TIMESTAMP_MS '2026-08-19 22:15:30.123' AS timestamp_ms, TIMESTAMP '2026-08-19 22:15:30.123456' AS timestamp_us, TIMESTAMP_NS '2026-08-19 22:15:30.123456789' AS timestamp_ns, TIMESTAMPTZ '2026-08-19 22:15:30.123456+00' AS summer, TIMESTAMPTZ '2026-01-19 22:15:30.123456+00' AS winter, 'epoch'::TIMESTAMP AS epoch_value, 'infinity'::TIMESTAMP AS positive_infinity, '-infinity'::TIMESTAMP AS negative_infinity, INTERVAL '1 year 2 months 3 days 04:05:06.123456' AS duration, DATE '1899-12-31' AS before_excel, DATE '9999-12-31' AS excel_max;", fallback);
        var result = HaybarnClient.ParseResult(output, null, 0);
        Equal("DATE", Column(result, "leap_date").Type, "date type inferred from Haybarn JSON");
        Equal("TIME", Column(result, "clock").Type, "time type inferred from Haybarn JSON");
        Equal("TIME WITH TIME ZONE", Column(result, "clock_tz").Type, "time zone type inferred from Haybarn JSON");
        Equal("TIMESTAMP_NS", Column(result, "timestamp_ns").Type, "nanosecond timestamp type inferred from Haybarn JSON");
        Equal("TIMESTAMP WITH TIME ZONE", Column(result, "summer").Type, "zoned timestamp type inferred from Haybarn JSON");
        var excel = WorkbookBridge.Values(result);
        Equal(new DateTime(2024, 2, 29), ExcelCell<DateTime>(result, excel, "leap_date"), "leap date sent to Excel");
        Equal(TimeSpan.Parse("23:59:59.123456", System.Globalization.CultureInfo.InvariantCulture).TotalDays, ExcelCell<double>(result, excel, "clock"), "time sent as Excel day fraction");
        Equal("23:59:59.123456-07", ExcelCell<string>(result, excel, "clock_tz"), "time with zone preserved as text");
        Equal(new DateTime(2026, 8, 19, 22, 15, 30), ExcelCell<DateTime>(result, excel, "timestamp_s"), "second timestamp round trip");
        Equal(new DateTime(2026, 8, 19, 22, 15, 30, 123), ExcelCell<DateTime>(result, excel, "timestamp_ms"), "millisecond timestamp round trip");
        Equal(new DateTime(2026, 8, 19, 22, 15, 30).AddTicks(1_234_560), ExcelCell<DateTime>(result, excel, "timestamp_us"), "microsecond timestamp round trip");
        Equal("2026-08-19 22:15:30.123456789", ExcelCell<string>(result, excel, "timestamp_ns"), "nanosecond timestamp preserved as exact text");
        Equal(new DateTime(2026, 8, 19, 18, 15, 30).AddTicks(1_234_560), ExcelCell<DateTime>(result, excel, "summer"), "summer offset converted to local Excel time");
        Equal(new DateTime(2026, 1, 19, 17, 15, 30).AddTicks(1_234_560), ExcelCell<DateTime>(result, excel, "winter"), "winter offset converted to local Excel time");
        Equal(new DateTime(1970, 1, 1), ExcelCell<DateTime>(result, excel, "epoch_value"), "epoch timestamp round trip");
        Equal("infinity", ExcelCell<string>(result, excel, "positive_infinity"), "positive temporal infinity preserved as text");
        Equal("-infinity", ExcelCell<string>(result, excel, "negative_infinity"), "negative temporal infinity preserved as text");
        Equal("1 year 2 months 3 days 04:05:06.123456", ExcelCell<string>(result, excel, "duration"), "interval preserved as text");
        Equal("1899-12-31", ExcelCell<string>(result, excel, "before_excel"), "pre-1900 date uses text fallback");
        Equal(new DateTime(9999, 12, 31), ExcelCell<DateTime>(result, excel, "excel_max"), "maximum Excel date round trip");
    }

    private static void NumericRoundTripTests()
    {
        const string fallback = "[{\"i8_min\":-128,\"i8_max\":127,\"i16_min\":-32768,\"i16_max\":32767,\"i32_min\":-2147483648,\"i32_max\":2147483647,\"i64_min\":-9223372036854775808,\"i64_max\":9223372036854775807,\"u8_max\":255,\"u16_max\":65535,\"u32_max\":4294967295,\"u64_max\":\"18446744073709551615\",\"i128_max\":\"170141183460469231731687303715884105727\",\"u128_max\":\"340282366920938463463374607431768211455\",\"safe_max\":9007199254740991,\"unsafe_pos\":9007199254740992,\"unsafe_neg\":-9007199254740992,\"dec38\":\"99999999999999999999999999999999999999\",\"dec_scale\":\"12345678901234567890.123456789012345678\",\"dec_tiny\":\".000000000000000001\",\"bignum_value\":\"12345678901234567890123456789012345678901234567890\",\"f32_max\":3.4028235e+38,\"f64_max\":1.7976931348623157e+308,\"f64_min\":5e-324,\"negative_zero\":-0.0,\"nan_value\":nan,\"positive_infinity\":inf,\"negative_infinity\":-inf}]";
        var sql = "SELECT (-128)::TINYINT AS i8_min, 127::TINYINT AS i8_max, (-32768)::SMALLINT AS i16_min, 32767::SMALLINT AS i16_max, (-2147483648)::INTEGER AS i32_min, 2147483647::INTEGER AS i32_max, (-9223372036854775807 - 1)::BIGINT AS i64_min, 9223372036854775807::BIGINT AS i64_max, 255::UTINYINT AS u8_max, 65535::USMALLINT AS u16_max, 4294967295::UINTEGER AS u32_max, 18446744073709551615::UBIGINT AS u64_max, 170141183460469231731687303715884105727::HUGEINT AS i128_max, 340282366920938463463374607431768211455::UHUGEINT AS u128_max, 9007199254740991::BIGINT AS safe_max, 9007199254740992::BIGINT AS unsafe_pos, (-9007199254740992)::BIGINT AS unsafe_neg, 99999999999999999999999999999999999999::DECIMAL(38,0) AS dec38, 12345678901234567890.123456789012345678::DECIMAL(38,18) AS dec_scale, 0.000000000000000001::DECIMAL(18,18) AS dec_tiny, '12345678901234567890123456789012345678901234567890'::BIGNUM AS bignum_value, 3.4028234663852886e38::FLOAT AS f32_max, 1.7976931348623157e308::DOUBLE AS f64_max, 4.9406564584124654e-324::DOUBLE AS f64_min, -0.0::DOUBLE AS negative_zero, 'NaN'::DOUBLE AS nan_value, 'Infinity'::DOUBLE AS positive_infinity, '-Infinity'::DOUBLE AS negative_infinity;";
        var result = HaybarnClient.ParseResult(RunHaybarnOrFallback(sql, fallback), null, 0);
        var values = WorkbookBridge.Values(result);
        Equal(-128L, ExcelCell<long>(result, values, "i8_min"), "TINYINT minimum");
        Equal(127L, ExcelCell<long>(result, values, "i8_max"), "TINYINT maximum");
        Equal(-32768L, ExcelCell<long>(result, values, "i16_min"), "SMALLINT minimum");
        Equal(32767L, ExcelCell<long>(result, values, "i16_max"), "SMALLINT maximum");
        Equal(-2147483648L, ExcelCell<long>(result, values, "i32_min"), "INTEGER minimum");
        Equal(2147483647L, ExcelCell<long>(result, values, "i32_max"), "INTEGER maximum");
        Equal("-9223372036854775808", ExcelCell<string>(result, values, "i64_min"), "BIGINT minimum preserved as text");
        Equal("9223372036854775807", ExcelCell<string>(result, values, "i64_max"), "BIGINT maximum preserved as text");
        Equal(255L, ExcelCell<long>(result, values, "u8_max"), "UTINYINT maximum");
        Equal(65535L, ExcelCell<long>(result, values, "u16_max"), "USMALLINT maximum");
        Equal(4294967295L, ExcelCell<long>(result, values, "u32_max"), "UINTEGER maximum");
        Equal("18446744073709551615", ExcelCell<string>(result, values, "u64_max"), "UBIGINT maximum");
        Equal("170141183460469231731687303715884105727", ExcelCell<string>(result, values, "i128_max"), "HUGEINT maximum");
        Equal("340282366920938463463374607431768211455", ExcelCell<string>(result, values, "u128_max"), "UHUGEINT maximum");
        Equal(9007199254740991L, ExcelCell<long>(result, values, "safe_max"), "largest exact Excel integer remains numeric");
        Equal("9007199254740992", ExcelCell<string>(result, values, "unsafe_pos"), "positive unsafe integer becomes text");
        Equal("-9007199254740992", ExcelCell<string>(result, values, "unsafe_neg"), "negative unsafe integer becomes text");
        Equal("99999999999999999999999999999999999999", ExcelCell<string>(result, values, "dec38"), "DECIMAL(38,0) exact text");
        Equal("12345678901234567890.123456789012345678", ExcelCell<string>(result, values, "dec_scale"), "scaled decimal exact text");
        Equal(".000000000000000001", ExcelCell<string>(result, values, "dec_tiny"), "small decimal exact text");
        Equal("12345678901234567890123456789012345678901234567890", ExcelCell<string>(result, values, "bignum_value"), "BIGNUM exact text");
        Equal("NaN", ExcelCell<string>(result, values, "nan_value"), "NaN normalized to Excel text");
        Equal("Infinity", ExcelCell<string>(result, values, "positive_infinity"), "positive infinity normalized to Excel text");
        Equal("-Infinity", ExcelCell<string>(result, values, "negative_infinity"), "negative infinity normalized to Excel text");
        Equal(double.MaxValue, ExcelCell<double>(result, values, "f64_max"), "DOUBLE maximum");
        Equal(double.Epsilon, ExcelCell<double>(result, values, "f64_min"), "DOUBLE subnormal minimum");
        Equal(0d, ExcelCell<double>(result, values, "negative_zero"), "negative zero normalizes to Excel numeric zero");
    }

    private static void TimeZoneEdgeTests()
    {
        const string fallback = "[{\"spring_before\":\"2026-03-08 01:59:59-05\",\"spring_after\":\"2026-03-08 03:00:00-04\",\"fall_first\":\"2026-11-01 01:30:00-04\",\"fall_second\":\"2026-11-01 01:30:00-05\",\"kathmandu\":\"2026-08-20 04:00:30\",\"excel_1900_before_bug\":\"1900-02-28\",\"excel_1900_after_bug\":\"1900-03-01\"}]";
        var sql = "SET TimeZone='America/New_York'; SELECT TIMESTAMPTZ '2026-03-08 06:59:59+00' AS spring_before, TIMESTAMPTZ '2026-03-08 07:00:00+00' AS spring_after, TIMESTAMPTZ '2026-11-01 05:30:00+00' AS fall_first, TIMESTAMPTZ '2026-11-01 06:30:00+00' AS fall_second, timezone('Asia/Kathmandu', TIMESTAMPTZ '2026-08-19 22:15:30+00') AS kathmandu, DATE '1900-02-28' AS excel_1900_before_bug, DATE '1900-03-01' AS excel_1900_after_bug;";
        var result = HaybarnClient.ParseResult(RunHaybarnOrFallback(sql, fallback), null, 0);
        Equal("2026-03-08 01:59:59-05", RawCell<string>(result, "spring_before"), "DST spring instant before gap");
        Equal("2026-03-08 03:00:00-04", RawCell<string>(result, "spring_after"), "DST spring instant after gap");
        Equal("2026-11-01 01:30:00-04", RawCell<string>(result, "fall_first"), "first ambiguous fall instant offset");
        Equal("2026-11-01 01:30:00-05", RawCell<string>(result, "fall_second"), "second ambiguous fall instant offset");
        var excel = WorkbookBridge.Values(result);
        Equal(new DateTime(2026, 3, 8, 1, 59, 59), ExcelCell<DateTime>(result, excel, "spring_before"), "DST before gap Excel wall time");
        Equal(new DateTime(2026, 3, 8, 3, 0, 0), ExcelCell<DateTime>(result, excel, "spring_after"), "DST after gap Excel wall time");
        Equal(new DateTime(2026, 11, 1, 1, 30, 0), ExcelCell<DateTime>(result, excel, "fall_first"), "first fall wall time");
        Equal(new DateTime(2026, 11, 1, 1, 30, 0), ExcelCell<DateTime>(result, excel, "fall_second"), "second fall wall time");
        Equal(new DateTime(2026, 8, 20, 4, 0, 30), ExcelCell<DateTime>(result, excel, "kathmandu"), "quarter-hour-offset timezone conversion");
        Equal(new DateTime(1900, 2, 28), ExcelCell<DateTime>(result, excel, "excel_1900_before_bug"), "last date before Excel serial 60 gap");
        Equal(new DateTime(1900, 3, 1), ExcelCell<DateTime>(result, excel, "excel_1900_after_bug"), "first date after Excel serial 60 gap");
    }

    private static void ExtendedNumericBoundaryTests()
    {
        const string fallback = "[{\"i128_min\":\"-170141183460469231731687303715884105728\",\"u8_min\":0,\"u16_min\":0,\"u32_min\":0,\"u64_min\":0,\"u128_min\":\"0\",\"bignum_negative\":\"-12345678901234567890123456789012345678901234567890\",\"dec4\":\"99.99\",\"dec9\":\"9999999.99\",\"dec18\":\"9999999999999999.99\",\"dec19\":\"99999999999999999.99\",\"dec_negative\":\"-99999999999999999999.999999999999999999\",\"f32_lowest\":-3.4028235e+38,\"f32_subnormal\":1e-45,\"f64_lowest\":-1.7976931348623157e+308}]";
        var sql = "SELECT (-170141183460469231731687303715884105727 - 1)::HUGEINT AS i128_min, 0::UTINYINT AS u8_min, 0::USMALLINT AS u16_min, 0::UINTEGER AS u32_min, 0::UBIGINT AS u64_min, 0::UHUGEINT AS u128_min, '-12345678901234567890123456789012345678901234567890'::BIGNUM AS bignum_negative, 99.99::DECIMAL(4,2) AS dec4, 9999999.99::DECIMAL(9,2) AS dec9, 9999999999999999.99::DECIMAL(18,2) AS dec18, 99999999999999999.99::DECIMAL(19,2) AS dec19, (-99999999999999999999.999999999999999999)::DECIMAL(38,18) AS dec_negative, (-3.4028234663852886e38)::FLOAT AS f32_lowest, 1.401298464324817e-45::FLOAT AS f32_subnormal, (-1.7976931348623157e308)::DOUBLE AS f64_lowest;";
        var result = HaybarnClient.ParseResult(RunHaybarnOrFallback(sql, fallback), null, 0);
        var excel = WorkbookBridge.Values(result);
        Equal("-170141183460469231731687303715884105728", ExcelCell<string>(result, excel, "i128_min"), "HUGEINT minimum");
        foreach (var name in new[] { "u8_min", "u16_min", "u32_min", "u64_min" }) Equal(0L, Convert.ToInt64(ExcelCell<object>(result, excel, name)), name + " unsigned minimum");
        Equal("0", Convert.ToString(ExcelCell<object>(result, excel, "u128_min")), "UHUGEINT minimum");
        Equal("-12345678901234567890123456789012345678901234567890", ExcelCell<string>(result, excel, "bignum_negative"), "negative BIGNUM exact text");
        Equal("99.99", ExcelCell<string>(result, excel, "dec4"), "DECIMAL width 4");
        Equal("9999999.99", ExcelCell<string>(result, excel, "dec9"), "DECIMAL width 9");
        Equal("9999999999999999.99", ExcelCell<string>(result, excel, "dec18"), "DECIMAL width 18");
        Equal("99999999999999999.99", ExcelCell<string>(result, excel, "dec19"), "DECIMAL width 19");
        Equal("-99999999999999999999.999999999999999999", ExcelCell<string>(result, excel, "dec_negative"), "negative DECIMAL(38,18)");
        True(ExcelCell<double>(result, excel, "f32_lowest") < -3.4e38, "FLOAT negative bound");
        True(ExcelCell<double>(result, excel, "f32_subnormal") > 0, "FLOAT positive subnormal");
        Equal(-double.MaxValue, ExcelCell<double>(result, excel, "f64_lowest"), "DOUBLE negative bound");
    }

    private sealed class FakeSession : IHaybarnSession
    {
        internal int Calls, Disposals, Active, Peak;
        public QueryResult Query(string sql, int? maxRows, System.Threading.CancellationToken cancellation = default)
        {
            var active = System.Threading.Interlocked.Increment(ref Active);
            Peak = Math.Max(Peak, active);
            try { System.Threading.Thread.Sleep(10); Calls++; if (sql == "fail") throw new InvalidOperationException("fixture failure"); return new QueryResult(); }
            finally { System.Threading.Interlocked.Decrement(ref Active); }
        }
        public void Dispose() { Disposals++; }
    }

    private static void SessionCacheTests()
    {
        var made = new List<FakeSession>();
        using var cache = new HaybarnSessionCache((_, cancellation) => { var value = new FakeSession(); made.Add(value); return value; });
        cache.Query("one", () => "setup", "ok", null);
        cache.Query("ONE", () => "setup", "ok", null);
        Equal(1, made.Count, "case-insensitive friendly name reuses session");
        System.Threading.Tasks.Parallel.For(0, 10, _ => cache.Query("one", () => "setup", "ok", null));
        Equal(1, made[0].Peak, "same session never executes concurrently");
        cache.Query("two", () => "setup", "ok", null);
        Equal(2, made.Count, "friendly identities remain isolated");
        cache.Query("one", () => "new credential or settings", "ok", null);
        Equal(3, made.Count, "changed attachment rebuilds session");
        Equal(1, made[0].Disposals, "changed session disposed once");
        Throws<InvalidOperationException>(() => cache.Query("one", () => "new credential or settings", "fail", null), "query failure not replayed");
        Equal(2, made[2].Calls, "failed SQL executed exactly once");
        Equal(1, made[2].Disposals, "failed session evicted");
        cache.Query("one", () => "new credential or settings", "ok", null);
        cache.Invalidate("one");
        Equal(1, made[3].Disposals, "signout invalidates session");
        cache.Query("one", () => "new credential or settings", "ok", null);
        cache.Dispose();
        True(made.All(value => value.Disposals == 1), "shutdown disposes all sessions exactly once");
        Throws<ObjectDisposedException>(() => cache.Query("one", () => "setup", "ok", null), "shutdown rejects new work");
        var attempts = 0;
        using var broken = new HaybarnSessionCache((_, cancellation) => { attempts++; throw new InvalidOperationException("attach failed"); });
        Throws<InvalidOperationException>(() => broken.Query("one", () => "setup", "ok", null), "failed attachment propagates");
        Throws<InvalidOperationException>(() => broken.Query("one", () => "setup", "ok", null), "failed attachment can be retried by a new request");
        Equal(2, attempts, "failed factory does not poison entry");
    }

    private static void WorkspaceTests()
    {
        var sales = new VgiConnection { Name = "workspace-sales", Catalog = "sales", Location = "https://sales.example.test" };
        var inventory = new VgiConnection { Name = "workspace-inventory", Catalog = "inventory", Location = "https://inventory.example.test" };
        ConnectionStore.Save(sales);
        ConnectionStore.Save(inventory, false);
        Equal(sales.Name, ConnectionStore.WorkspaceName(), "existing default initializes workspace");
        var bridgeResult = WorkbenchBridge.Invoke("connections.workspace", new JObject { ["members"] = new JArray(sales.Name, inventory.Name) }).GetAwaiter().GetResult();
        var listed = JArray.FromObject(bridgeResult!);
        Equal(1, listed.Count(item => item.Value<bool>("IsWorkspaceSelected")), "bridge exposes one shared workspace selection");
        Equal(sales.Name, listed.Single(item => item.Value<bool>("IsDefault")).Value<string>("Name"), "bridge preserves the legacy default separately");
        var combined = ConnectionStore.Resolve(ConnectionStore.WorkspaceName());
        Equal(true, combined.IsWorkspaceProfile, "workspace creates a retained connection set");
        Equal(2, ConnectionStore.ResolveAttachments(combined).Count, "workspace attaches both catalogs");
        Equal(sales.Name, ConnectionStore.DefaultName(), "workspace selection preserves legacy formula default");
        var oldName = combined.Name;
        ConnectionStore.SetWorkspace(new[] { inventory.Name, sales.Name });
        Equal(inventory.Name, ConnectionStore.Resolve(ConnectionStore.WorkspaceName()).Members[0], "workspace default changes unqualified SQL catalog");
        Equal(sales.Name, ConnectionStore.Resolve(oldName).Members[0], "existing Power Query identity keeps original default");
        ConnectionStore.SetWorkspace(new[] { sales.Name.ToUpperInvariant(), inventory.Name });
        Equal(oldName, ConnectionStore.WorkspaceName(), "same catalog set reuses its identity");
        ConnectionStore.SetWorkspace(new[] { inventory.Name });
        Equal(inventory.Name, ConnectionStore.WorkspaceName(), "single-catalog workspace uses saved identity");
        Equal(2, ConnectionStore.ResolveAttachments(ConnectionStore.Resolve(oldName)).Count, "removing a workspace catalog preserves saved refresh attachments");
        Throws<ArgumentException>(() => ConnectionStore.SetWorkspace(Array.Empty<string>()), "empty workspace rejected");
        Throws<InvalidOperationException>(() => ConnectionStore.SetWorkspace(new[] { "missing" }), "missing workspace catalog rejected");
        Throws<ArgumentException>(() => ConnectionStore.SetWorkspace(new[] { inventory.Name, inventory.Name }), "duplicate workspace catalog rejected");
        Throws<ArgumentException>(() => ConnectionStore.SetWorkspace(new[] { oldName }), "nested workspace profile rejected");
        Equal(inventory.Name, ConnectionStore.WorkspaceName(), "failed update preserves workspace");
        Throws<ArgumentException>(() => ConnectionStore.Save(new VgiConnection { Name = oldName, Catalog = "replacement", Location = "https://example.test" }, false), "retained query set cannot be overwritten");
        Throws<InvalidOperationException>(() => ConnectionStore.Remove(oldName), "retained query set cannot be deleted");
        var conflicting = new VgiConnection { Name = "workspace-conflict", Catalog = "inventory", Location = "https://other.example.test" };
        ConnectionStore.Save(conflicting, false);
        Throws<ArgumentException>(() => ConnectionStore.SetWorkspace(new[] { inventory.Name, conflicting.Name }), "ambiguous catalog aliases rejected");
        var setup = HaybarnClient.BuildSessionScript(ConnectionStore.ResolveAttachments(ConnectionStore.Resolve(oldName)), "SELECT 1", timeZone: "UTC");
        True(setup.Contains("AS \"sales\""), "workspace setup attaches sales");
        True(setup.Contains("AS \"inventory\""), "workspace setup attaches inventory");
    }

    private static void ProfileTests()
    {
        var weather = new VgiConnection { Name = "profile-weather", Catalog = "weather", Location = "https://weather.example.test" };
        var earthquakes = new VgiConnection { Name = "profile-earthquakes", Catalog = "earthquakes", Location = "https://earthquakes.example.test" };
        var profile = new VgiConnection { Name = "profile-research", Members = new[] { weather.Name.ToUpperInvariant(), earthquakes.Name } };
        var saved = new[] { weather, earthquakes, profile };
        var attachments = ConnectionStore.ResolveAttachments(profile, saved);
        Equal(weather, attachments[0], "profile resolves members case-insensitively in default order");
        Equal(2, attachments.Count, "all profile members resolved");
        var setup = HaybarnClient.BuildSessionScript(attachments, "SELECT 1", "UTC");
        True(setup.Contains("AS \"weather\"") && setup.Contains("AS \"earthquakes\"") && setup.Contains("USE \"weather\""), "profile attaches both and selects the first catalog");
        Throws<InvalidOperationException>(() => ConnectionStore.ResolveAttachments(profile, new[] { weather }), "missing profile member fails closed");
        earthquakes.Catalog = "WEATHER";
        Throws<ArgumentException>(() => ConnectionStore.ResolveAttachments(profile, saved), "duplicate catalog aliases rejected");
        earthquakes.Catalog = "earthquakes";
        earthquakes.Members = new[] { weather.Name };
        Throws<ArgumentException>(() => ConnectionStore.ResolveAttachments(profile, saved), "nested profile rejected");
        earthquakes.Members = Array.Empty<string>();
        profile.Members = new[] { weather.Name, weather.Name.ToUpperInvariant() };
        Throws<ArgumentException>(() => ConnectionStore.ResolveAttachments(profile, saved), "duplicate member rejected");
        profile.Members = new[] { weather.Name, earthquakes.Name };
        var identities = attachments.Select(member => new JObject { ["contract_version"] = 2, ["connection_name"] = profile.Name, ["member_name"] = member.Name, ["catalog_alias"] = member.Catalog, ["location"] = member.Location, ["authentication"] = member.Authentication, ["attach_options"] = "{}" }).ToArray();
        PowerQueryBridge.ValidateIdentities(profile, attachments, Enumerable.Reverse(identities).ToArray());
        Throws<InvalidOperationException>(() => PowerQueryBridge.ValidateIdentities(profile, attachments, identities.Take(1).ToArray()), "missing driver attachment rejected");
        foreach (var field in new[] { "connection_name", "member_name", "catalog_alias", "location", "authentication", "attach_options" })
        {
            var wrong = identities.Select(row => (JObject)row.DeepClone()).ToArray();
            wrong[1][field] = field == "attach_options" ? "{\"changed\":true}" : "different";
            Throws<InvalidOperationException>(() => PowerQueryBridge.ValidateIdentities(profile, attachments, wrong), "second attachment mismatch rejected: " + field);
        }
        ConnectionStore.Save(weather); ConnectionStore.Save(earthquakes); ConnectionStore.Save(profile);
        try { Throws<InvalidOperationException>(() => ConnectionStore.Remove(weather.Name), "referenced member cannot be removed"); }
        finally { ConnectionStore.Remove(profile.Name); ConnectionStore.Remove(weather.Name); ConnectionStore.Remove(earthquakes.Name); }
    }

    private static void ConnectionProbeTests()
    {
        var anonymous = ConnectionProbe.CatalogScript("https://example.test", null);
        True(anonymous.Contains("vgi_oauth_enabled=false"), "discovery cannot start engine sign-in");
        True(anonymous.Contains("oauth_cache := 'none'"), "discovery does not persist engine credentials");
        var refresh = ConnectionProbe.CatalogScript("https://example.test", new OAuthAttachCredential("oauth_refresh_token", "test'token"));
        True(refresh.Contains("vgi_oauth_enabled=true"), "discovery enables refresh for saved OAuth credentials");
        True(refresh.Contains("oauth_refresh_token := 'test''token'"), "discovery safely quotes the in-memory refresh credential");
        var bearer = ConnectionProbe.CatalogScript("https://example.test", new OAuthAttachCredential("bearer_token", "test-only"));
        True(bearer.Contains("vgi_oauth_enabled=false") && bearer.Contains("bearer_token := 'test-only'"), "bearer discovery uses the saved credential without engine sign-in");
        using var finished = new System.Threading.ManualResetEventSlim();
        var clock = Stopwatch.StartNew();
        Throws<TimeoutException>(() => ConnectionProbe.Run<int>(token => {
            try { token.WaitHandle.WaitOne(); token.ThrowIfCancellationRequested(); return 0; }
            finally { finished.Set(); }
        }, 30).GetAwaiter().GetResult(), "draft deadline cancels blocked work");
        True(clock.Elapsed < TimeSpan.FromSeconds(3), "deadline releases caller promptly");
        True(finished.Wait(3000), "timed-out work observes cancellation");
        // The finally block releases the gate just after the fixture signals completion.
        System.Threading.Thread.Sleep(20);
        Equal(42, ConnectionProbe.Run(_ => 42).GetAwaiter().GetResult(), "retry succeeds after cancellation");
        Throws<ArgumentException>(() => ConnectionProbe.Run<int>(_ => throw new ArgumentException("fixture")).GetAwaiter().GetResult(), "probe errors propagate");
        Equal(7, ConnectionProbe.Run(_ => 7).GetAwaiter().GetResult(), "failure releases the probe slot");
        Throws<ArgumentException>(() => ConnectionProbe.Catalogs("http://example.test").GetAwaiter().GetResult(), "discovery requires HTTPS at native boundary");
        var original = new VgiConnection { Name = "duplicate-test", Catalog = "original", Location = "https://example.test" };
        ConnectionStore.Save(original, false);
        try {
            var draft = new VgiConnection { Name = "DUPLICATE-TEST", Catalog = "changed", Location = "https://other.example.test" };
            Throws<ArgumentException>(() => WorkbenchBridge.Invoke("connections.save", new JObject { ["connection"] = JObject.FromObject(draft, WorkbenchBridge.Serializer), ["originalName"] = "" }).GetAwaiter().GetResult(), "bridge rejects a duplicate new name");
            Equal("original", ConnectionStore.Resolve(original.Name).Catalog, "duplicate save preserves stored connection");
        } finally { ConnectionStore.Remove(original.Name); }
    }

    private static void NativeConnectionProbeTests()
    {
        var endpoint = Environment.GetEnvironmentVariable("CUPOLA_LIVE_VGI_ENDPOINT") ?? "https://vgi-open-meteo.rusty-bb6.workers.dev";
        var saved = new VgiConnection { Name = "probe-saved", Catalog = "open_meteo", Location = endpoint };
        ConnectionStore.Save(saved);
        var before = Newtonsoft.Json.JsonConvert.SerializeObject(ConnectionStore.List());
        var defaultBefore = ConnectionStore.DefaultName();
        var catalogs = (string[])WorkbenchBridge.Invoke("connections.catalogs", new JObject { ["location"] = endpoint }).GetAwaiter().GetResult()!;
        True(catalogs.Contains("open_meteo"), "native discovery lists catalog before ATTACH");
        var draft = new VgiConnection { Name = "unsaved-probe", Catalog = "open_meteo", Location = endpoint };
        WorkbenchBridge.Invoke("connections.test", new JObject { ["connection"] = JObject.FromObject(draft, WorkbenchBridge.Serializer) }).GetAwaiter().GetResult();
        Equal(before, Newtonsoft.Json.JsonConvert.SerializeObject(ConnectionStore.List()), "successful draft test never saves a connection");
        draft.Name = saved.Name;
        draft.Location = "https://vgi-open-meteo.rusty-bb6.workers.de";
        var clock = Stopwatch.StartNew();
        try { WorkbenchBridge.Invoke("connections.test", new JObject { ["connection"] = JObject.FromObject(draft, WorkbenchBridge.Serializer) }).GetAwaiter().GetResult(); throw new Exception("Misspelled hostname unexpectedly succeeded"); }
        catch (Exception error) when (error is InvalidOperationException || error is TimeoutException) { }
        True(clock.Elapsed < TimeSpan.FromSeconds(23), "native bad host returns within deadline");
        Equal(before, Newtonsoft.Json.JsonConvert.SerializeObject(ConnectionStore.List()), "failed edit test preserves saved connection");
        Equal(defaultBefore, ConnectionStore.DefaultName(), "testing preserves default connection");
        Equal("https://vgi-open-meteo.rusty-bb6.workers.de", draft.Location, "native probe never corrects hostnames");
        // A real failed native request must release the disposable session for retry.
        draft.Location = endpoint;
        Equal(1, ConnectionProbe.Test(draft).GetAwaiter().GetResult().RowCount, "native probe retry succeeds");
        ConnectionStore.Remove(saved.Name);
        Console.WriteLine("PASS: native URL discovery, draft isolation, bad-host deadline, and retry");
    }

    private static void NativeSessionTests()
    {
        NativeConnectionProbeTests();
        Console.WriteLine("Native session: opening engine");
        using (var native = new NativeHaybarnSession("SET TimeZone='America/New_York'"))
        {
            using (var cancel = new System.Threading.CancellationTokenSource())
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                cancel.CancelAfter(250);
                Throws<OperationCanceledException>(() => native.Query("SELECT sum(sin(i)) FROM range(100000000000) t(i)", 1, cancel.Token), "native running query is interrupted");
                True(watch.Elapsed < TimeSpan.FromSeconds(10), "native cancellation is prompt");
            }
            Equal(42L, native.Query("SELECT 42 AS after_cancel", 1).Rows[0][0], "native connection remains usable after cancellation");
            Console.WriteLine("Native session: scalar conversion");
            var result = native.Query("SELECT 99.99::DECIMAL(18,2) AS amount, 99999999999999.99::DECIMAL(38,2) AS precise, '001200' AS code, 'Ω雪' AS unicode, 9223372036854775807::BIGINT AS large, NULL::INTEGER AS missing, true AS flag", null);
            Equal(99.99d, result.Rows[0][0], "native decimal remains numeric");
            Equal("DECIMAL(18,2)", result.Columns[0].Type, "native logical decimal type");
            Equal("99999999999999.99", result.Rows[0][1], "native precise decimal remains text");
            Equal("001200", result.Rows[0][2], "native numeric-looking strings preserved");
            Equal("Ω雪", result.Rows[0][3], "native UTF8 round trip");
            Equal("9223372036854775807", result.Rows[0][4], "native unsafe integer preserved");
            True(result.Rows[0][5] is null, "native SQL null");
            Equal(true, result.Rows[0][6], "native boolean");
            Console.WriteLine("Native session: empty result");
            var typed = native.Query("SELECT TIMESTAMPTZ '2026-08-19 22:15:30.123456+00' AS summer, TIMESTAMPTZ '2026-01-19 22:15:30.123456+00' AS winter, TIMESTAMP_NS '2026-08-19 22:15:30.123456789' AS nanos, [1,NULL,3] AS items, {'a':42,'b':'雪'} AS record, MAP {'a':1,'b':2} AS mapping, '10101'::BIT AS bits, '1234567890123456789012345678901234567890'::BIGNUM AS big, UUID '123e4567-e89b-12d3-a456-426614174000' AS id", null);
            Equal("2026-08-19 18:15:30.123456-04:00", typed.Rows[0][0], "native summer time zone");
            Equal("2026-01-19 17:15:30.123456-05:00", typed.Rows[0][1], "native winter time zone");
            Equal("2026-08-19 22:15:30.123456789", typed.Rows[0][2], "native nanoseconds preserved");
            Equal("[1,null,3]", typed.Rows[0][3], "native nested list JSON");
            Equal("{\"a\":42,\"b\":\"雪\"}", typed.Rows[0][4], "native struct JSON");
            Equal("{\"a\":1,\"b\":2}", typed.Rows[0][5], "native map JSON");
            Equal("10101", typed.Rows[0][6], "native bit string");
            Equal("1234567890123456789012345678901234567890", typed.Rows[0][7], "native arbitrary integer");
            Equal("123e4567-e89b-12d3-a456-426614174000", typed.Rows[0][8], "native UUID");
            var chunks = native.Query("SELECT range AS n, CASE WHEN range%65=0 THEN NULL ELSE 'long Unicode string 雪 ' || range::VARCHAR END AS text FROM range(5000)", null);
            Equal(5000, chunks.RowCount, "multiple chunks returned");
            Equal(4999L, chunks.Rows[4999][0], "last chunk numeric value");
            True(chunks.Rows[65][1] is null, "validity mask crosses 64-bit boundary");
            Equal("long Unicode string 雪 4999", chunks.Rows[4999][1], "heap string across chunks");
            var empty = native.Query("SELECT NULL::DECIMAL(18,2) AS amount WHERE false", null);
            Equal(0, empty.RowCount, "native empty result");
            Equal("DECIMAL(18,2)", empty.Columns[0].Type, "empty result retains schema");
            var limited = native.Query("SELECT range AS n FROM range(5)", 2);
            Equal(5, limited.RowCount, "full row count retained"); Equal(2, limited.Rows.Length, "native preview bounded"); True(limited.Truncated, "native truncation flagged");
            native.Query("CREATE TEMP TABLE reuse_proof AS SELECT 42 AS n", null);
            Equal(42L, native.Query("SELECT n FROM reuse_proof", null).Rows[0][0], "native session persists state");
            Equal("a\0b", native.Query("SELECT 'a' || chr(0) || 'b'", null).Rows[0][0], "embedded NUL preserved");
        }
        var connection = new VgiConnection { Name = "native-session-fixture", Catalog = "open_meteo", Location = Environment.GetEnvironmentVariable("CUPOLA_LIVE_VGI_ENDPOINT") ?? "https://vgi-open-meteo.rusty-bb6.workers.dev" };
        ConnectionStore.Save(connection);
        var client = new HaybarnClient();
        var timings = new List<double>();
        for (var i = 0; i < 4; i++)
        {
            var clock = Stopwatch.StartNew();
            var value = client.QueryResult("SELECT open_meteo.main.weather_code_text(0) AS label", connection.Name);
            Equal("Clear sky", value.Rows[0][0], "native VGI HTTPS result");
            timings.Add(clock.Elapsed.TotalMilliseconds);
        }
        client.QueryResult("CREATE TEMP TABLE session_proof AS SELECT 123 AS n", connection.Name);
        Equal(123L, client.QueryResult("SELECT n FROM session_proof", connection.Name).Rows[0][0], "client reuses native attachment");
        foreach (var queryMethod in new[] { "query.editor", "query.agent" })
        {
            var queryId = Guid.NewGuid().ToString();
            var running = WorkbenchBridge.Invoke(queryMethod, new JObject { ["sql"] = "SELECT sum(sin(i)) FROM range(100000000000) t(i)", ["connection"] = connection.Name, ["queryId"] = queryId });
            System.Threading.Thread.Sleep(250);
            Equal(false, WorkbenchBridge.Invoke("query.cancel", new JObject { ["queryId"] = Guid.NewGuid().ToString() }).GetAwaiter().GetResult(), "unrelated cancellation cannot target running query");
            var queuedId = Guid.NewGuid().ToString();
            var queued = WorkbenchBridge.Invoke(queryMethod, new JObject { ["sql"] = "SELECT 1", ["connection"] = connection.Name, ["queryId"] = queuedId });
            Equal(true, WorkbenchBridge.Invoke("query.cancel", new JObject { ["queryId"] = queuedId }).GetAwaiter().GetResult(), "queued cancellation accepted");
            True(queued.Wait(5000), "queued cancellation does not wait for active query");
            True(queued.Result is null, "queued query reports cancellation");
            True(!running.IsCompleted, "queued cancel leaves running query alone");
            Equal(true, WorkbenchBridge.Invoke("query.cancel", new JObject { ["queryId"] = queryId }).GetAwaiter().GetResult(), "active cancellation accepted");
            True(running.Wait(10000), "bridge cancellation settles promptly");
            True(running.Result is null, "bridge confirms query cancelled");
            Equal(false, WorkbenchBridge.Invoke("query.cancel", new JObject { ["queryId"] = queryId }).GetAwaiter().GetResult(), "completed query cancellation is harmless");
            Equal(123L, client.QueryResult("SELECT n FROM session_proof", connection.Name).Rows[0][0], "cancellation preserves cached native session state");
        }
        connection.AttachOptions["nonexistent_option_fixture"] = true;
        ConnectionStore.Save(connection);
        Throws<InvalidOperationException>(() => client.QueryResult("SELECT 1", connection.Name), "changed options invalidate existing attachment");
        connection.AttachOptions.Clear(); ConnectionStore.Save(connection);
        Throws<InvalidOperationException>(() => client.QueryResult("SELECT n FROM session_proof", connection.Name), "reconnected session cannot reuse old temporary state");
        Console.WriteLine("Native session HTTPS timings ms (cold then warm): " + string.Join(", ", timings.Select(value => value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture))));
        var earthquakes = new VgiConnection { Name = "native-earthquakes-fixture", Catalog = "earthquakes", Location = "https://vgi-earthquakes.rusty-bb6.workers.dev" };
        var profile = new VgiConnection { Name = "native-profile-fixture", Members = new[] { connection.Name, earthquakes.Name } };
        ConnectionStore.Save(earthquakes); ConnectionStore.Save(profile);
        var combined = client.QueryResult("SELECT open_meteo.main.weather_code_text(0) AS weather, count(*) AS earthquakes FROM (SELECT * FROM earthquakes.main.recent LIMIT 1)", profile.Name);
        Equal("Clear sky", combined.Rows[0][0], "native query uses weather and earthquakes together");
        Equal(1L, combined.Rows[0][1], "native second catalog returns live data");
        Equal("open_meteo", client.QueryResult("SELECT current_catalog()", profile.Name).Rows[0][0], "profile default catalog");
        client.QueryResult("CREATE TEMP TABLE profile_proof AS SELECT 1", profile.Name);
        earthquakes.AttachOptions["nonexistent_option_fixture"] = true; ConnectionStore.Save(earthquakes);
        Throws<InvalidOperationException>(() => client.QueryResult("SELECT 1", profile.Name), "member changes invalidate profile attachment");
        earthquakes.AttachOptions.Clear(); ConnectionStore.Save(earthquakes);
        Throws<InvalidOperationException>(() => client.QueryResult("SELECT * FROM profile_proof", profile.Name), "profile reconnect cannot reuse old temporary state");
        Equal("Clear sky", client.QueryResult("SELECT open_meteo.main.weather_code_text(0) FROM earthquakes.main.recent LIMIT 1", profile.Name).Rows[0][0], "profile recovers after member settings are restored");
        HaybarnSessions.Cache.Dispose();
        Console.WriteLine("PASS: native session lifecycle, values, and HTTPS reuse");
    }

    private static void AccountingDecimalTests()
    {
        const string payload = "[{\"column_name\":\"amount\",\"column_type\":\"DECIMAL(18,2)\"},{\"column_name\":\"credit\",\"column_type\":\"DECIMAL(18,2)\"},{\"column_name\":\"unsafe_amount\",\"column_type\":\"DECIMAL(38,2)\"},{\"column_name\":\"account_code\",\"column_type\":\"VARCHAR\"}]\n[{\"amount\":\"99.99\",\"credit\":\"-1234.50\",\"unsafe_amount\":\"99999999999999.99\",\"account_code\":\"001200\"}]";
        const string sql = "SELECT 99.99::DECIMAL(18,2) AS amount, (-1234.50)::DECIMAL(18,2) AS credit, 99999999999999.99::DECIMAL(38,2) AS unsafe_amount, '001200'::VARCHAR AS account_code";
        var result = HaybarnClient.ParseResult(RunHaybarnOrFallback(HaybarnClient.AddDescribePrelude(sql), payload), null, 0);
        Equal("DECIMAL(18,2)", Column(result, "amount").Type, "declared decimal type retained");
        var values = WorkbookBridge.Values(result);
        Equal(99.99d, ExcelCell<double>(result, values, "amount"), "ordinary debit remains numeric");
        Equal(-1234.5d, ExcelCell<double>(result, values, "credit"), "ordinary credit remains numeric");
        Equal("99999999999999.99", ExcelCell<string>(result, values, "unsafe_amount"), "unsafe monetary precision remains exact text");
        Equal("001200", ExcelCell<string>(result, values, "account_code"), "account code retains leading zeroes");
        Equal("#,##0.00", WorkbookBridge.NumberFormat("DECIMAL(18,2)", result.Rows, 0), "monetary number format");
        Equal("@", WorkbookBridge.NumberFormat("DECIMAL(38,2)", result.Rows, 2), "unsafe decimal text format");
        True(HaybarnClient.AddDescribePrelude("SELECT ';' AS marker").StartsWith("DESCRIBE "), "type prelude ignores semicolons inside literals");
        Equal("SELECT 1; SELECT 2", HaybarnClient.AddDescribePrelude("SELECT 1; SELECT 2"), "multi-statement SQL skips type prelude");
    }

    private static string RunHaybarnOrFallback(string sql, string fallback)
    {
        var engine = Environment.GetEnvironmentVariable("VGI_HAYBARN_PATH");
        if (string.IsNullOrWhiteSpace(engine) || !File.Exists(engine))
        {
            Console.WriteLine("SKIP: VGI_HAYBARN_PATH is not set; validating a captured Haybarn compatibility payload.");
            return fallback;
        }
        var start = new ProcessStartInfo(engine, "-json")
        {
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            UseShellExecute = false, CreateNoWindow = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
        };
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Unable to start Haybarn for the compatibility test.");
        process.StandardInput.WriteLine(sql);
        process.StandardInput.Close();
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) throw new InvalidOperationException("Local compatibility query failed: " + error);
        return output;
    }

    private static QueryColumn Column(QueryResult result, string name) => result.Columns.Single(column => column.Name == name);
    private static T RawCell<T>(QueryResult result, string name) => (T)result.Rows[0][Array.FindIndex(result.Columns, column => column.Name == name)]!;
    private static T ExcelCell<T>(QueryResult result, object[,] values, string name) => (T)values[1, Array.FindIndex(result.Columns, column => column.Name == name)];

    private static void WebWorkbenchTests()
    {
        var assets = Environment.GetEnvironmentVariable("VGI_EXCEL_WEB_ASSETS_PATH");
        if (!Environment.UserInteractive || string.IsNullOrWhiteSpace(assets) || !Directory.Exists(assets)) return;
        foreach (var connection in ConnectionStore.List()) ConnectionStore.Remove(connection.Name);
        using var form = WebWorkbenchForm.Create(3);
        True(form is WebWorkbenchForm, "web assets should select the embedded Workbench");
        True(form.ShowIcon && form.Icon is not null && form.Icon.Size == new System.Drawing.Size(64, 64), "main window uses the embedded Cupola icon before browser startup");
        using (var fallback = new NativeWorkbenchForm())
            True(fallback.ShowIcon && fallback.Icon is not null && fallback.Icon.Size == form.Icon!.Size, "native fallback uses the Cupola window icon");
        var opening = form.Controls.Find("openingPanel", false).Single();
        True(opening.Controls.OfType<System.Windows.Forms.PictureBox>().Single().Image is not null, "startup mark is embedded in the packed assembly");
        True(opening.Controls.OfType<System.Windows.Forms.Label>().Single().Text == "Opening Cupola…", "startup text is available before browser initialization");
        form.Show();
        True(opening.Visible, "native startup cover is visible immediately");
        for (var attempt = 0; attempt < 200 && opening.Visible; attempt++)
        {
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(100);
        }
        Equal("Ready", WebWorkbenchForm.LastStatus, "WebView2 navigation status");
        True(!opening.Visible, "paint handshake removes native startup cover");
        var web = form.Controls.OfType<Microsoft.Web.WebView2.WinForms.WebView2>().Single();
        string ReadPage()
        {
            var task = web.CoreWebView2.ExecuteScriptAsync("document.body.innerText");
            for (var attempt = 0; attempt < 100 && !task.IsCompleted; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50); }
            True(task.IsCompleted, "WebView page inspection completes");
            return Newtonsoft.Json.JsonConvert.DeserializeObject<string>(task.GetAwaiter().GetResult()) ?? "";
        }
        string InspectScript(string script)
        {
            var task = web.CoreWebView2.ExecuteScriptAsync(script);
            for (var attempt = 0; attempt < 100 && !task.IsCompleted; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50); }
            True(task.IsCompleted, "WebView motion inspection completes");
            return task.GetAwaiter().GetResult();
        }
        var reducedMotion = InspectScript("matchMedia('(prefers-reduced-motion: reduce)').matches") == "true";
        InspectScript("(() => { const box = document.createElement('div'); box.id = 'motion-probe'; box.className = 'catalog-loading'; box.innerHTML = '<svg class=busy-spinner width=18 height=18><circle cx=9 cy=9 r=7 /></svg>'; document.body.append(box); })()");
        var firstTransform = InspectScript("getComputedStyle(document.querySelector('#motion-probe svg')).transform");
        for (var tick = 0; tick < 5; tick++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50); }
        var nextTransform = InspectScript("getComputedStyle(document.querySelector('#motion-probe svg')).transform");
        True(reducedMotion ? firstTransform == nextTransform : firstTransform != nextTransform, "busy spinner follows actual WebView motion preference");
        InspectScript("document.querySelector('#motion-probe').remove()");
        Console.WriteLine("WebView motion preference: " + (reducedMotion ? "reduced (static busy icon)" : "normal (spinner animation verified)"));
        string page = "";
        for (var attempt = 0; attempt < 100; attempt++)
        {
            page = ReadPage();
            if (page.Contains("New connection")) break;
            System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50);
        }
        True(page.Contains("New connection") && !page.Contains("Cupola needs to restart"), "empty registry opens native-hosted Connections without a render crash");
        ConnectionStore.Save(new VgiConnection { Name = "persisted-webview-test", Catalog = "sample", Location = "https://example.test" });
        web.CoreWebView2.Reload();
        for (var attempt = 0; attempt < 100; attempt++)
        {
            System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50);
            page = ReadPage();
            if (page.Contains("persisted-webview-test")) break;
        }
        True(page.Contains("persisted-webview-test") && !page.Contains("Cupola needs to restart"), "saved native connection survives WebView reload");
        var popupRequest = new JObject { ["id"] = 9876, ["method"] = "results.open", ["params"] = new JObject {
            ["title"] = "Window fixture", ["result"] = JObject.FromObject(new QueryResult { Columns = new[] { new QueryColumn { Name = "value", Type = "VARCHAR" } }, Rows = new[] { new object?[] { "<b>literal result</b>" } }, RowCount = 1 }, WorkbenchBridge.Serializer)
        } };
        _ = web.CoreWebView2.ExecuteScriptAsync("window.chrome.webview.postMessage(" + popupRequest.ToString(Newtonsoft.Json.Formatting.None) + ")");
        for (var attempt = 0; attempt < 100 && !form.OwnedForms.Any(); attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50); }
        True(form.OwnedForms.Length == 1, "result opens in separate native window");
        var viewer = form.OwnedForms.Single();
        True(viewer.ShowIcon && viewer.Icon is not null && viewer.Icon.Size == form.Icon!.Size, "results window uses the same embedded Cupola icon");
        var viewerWeb = viewer.Controls.OfType<Microsoft.Web.WebView2.WinForms.WebView2>().Single();
        string ViewerScript(string script)
        {
            var task = viewerWeb.CoreWebView2.ExecuteScriptAsync(script);
            for (var attempt = 0; attempt < 100 && !task.IsCompleted; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50); }
            return task.GetAwaiter().GetResult();
        }
        for (var attempt = 0; attempt < 200; attempt++)
        {
            System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(50);
            if (viewerWeb.CoreWebView2 is not null && ViewerScript("document.body.innerText").Contains("literal result")) break;
        }
        True(ViewerScript("document.body.innerText").Contains("literal result"), "native result snapshot reaches viewer");
        True(!viewer.Controls.Find("openingPanel", false).Single().Visible, "results paint handshake removes native startup cover");
        Equal("0", ViewerScript("document.querySelectorAll('td b').length"), "native result cells render as text");
        ViewerScript("window.chrome.webview.addEventListener('message', e => { if(e.data.id === 9878) window.deniedResultAction = !!e.data.error; }); window.chrome.webview.postMessage({id:9878,method:'connections.list',params:{}})");
        for (var attempt = 0; attempt < 100 && ViewerScript("window.deniedResultAction") != "true"; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(20); }
        Equal("true", ViewerScript("window.deniedResultAction"), "results window cannot access connection bridge");
        ViewerScript("window.chrome.webview.postMessage({id:9877,method:'results.maximize',params:{}})");
        for (var attempt = 0; attempt < 100 && viewer.WindowState != System.Windows.Forms.FormWindowState.Maximized; attempt++) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(20); }
        Equal(System.Windows.Forms.FormWindowState.Maximized, viewer.WindowState, "results window can maximize");
        viewer.Close();
        True(ReadPage().Contains("persisted-webview-test"), "closing viewer preserves parent workbench");
        // Windows caps a large window at the remote desktop's available height.
        // Start small so the growth assertion tests resizing, not that OS limit.
        form.Size = form.MinimumSize;
        System.Windows.Forms.Application.DoEvents();
        var grip = form.Controls.OfType<System.Windows.Forms.StatusStrip>().Single();
        True(grip.SizingGrip && grip.Visible, "native resize grip remains visible outside WebView");
        var corner = grip.PointToScreen(new System.Drawing.Point(grip.ClientSize.Width - 3, grip.ClientSize.Height - 3));
        var hit = SendMessage(grip.Handle, 0x0084, IntPtr.Zero, new IntPtr((corner.Y << 16) | (corner.X & 0xffff)));
        Equal(new IntPtr(17), hit, "resize grip hit test selects both axes (HTBOTTOMRIGHT)");
        var before = form.Size;
        form.Size = new System.Drawing.Size(before.Width + 60, before.Height + 40);
        Equal(before.Width + 60, form.Width, "window can grow horizontally");
        Equal(before.Height + 40, form.Height, "window can grow vertically");
        typeof(WebWorkbenchForm).GetMethod("StartupFailed", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!.Invoke(form, new object[] { new TimeoutException("Close this window and try again.") });
        True(web.IsDisposed, "startup failure disposes the browser and releases its profile");
        True(form.Controls.OfType<System.Windows.Forms.FlowLayoutPanel>().Single().Controls.OfType<System.Windows.Forms.Button>().Any(button => button.Text == "Open native Cupola"), "startup failure offers a native recovery action");
        form.Close();
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);

    private static void Equal<T>(T expected, T actual, string label)
    {
        if (!Equals(expected, actual)) throw new InvalidOperationException($"{label}: expected {expected}, got {actual}");
    }

    private static void True(bool value, string label)
    {
        if (!value) throw new InvalidOperationException(label);
    }

    private static void Throws<T>(Action action, string label) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException(label + ": expected " + typeof(T).Name);
    }
}
