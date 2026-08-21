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
    private static int Main()
    {
        var root = Path.Combine(Path.GetTempPath(), "vgi-excel-tests-" + Guid.NewGuid().ToString("N"));
        Environment.SetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME", root);
        Environment.SetEnvironmentVariable("VGI_EXCEL_ANTHROPIC_CREDENTIAL_TARGET", "QueryFarm/VgiExcel/Tests/" + Guid.NewGuid().ToString("N"));
        try
        {
            ConnectionPolicyTests(root);
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
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
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

        var success = OAuthCallbackPage.Success("Nearwater <Production>", "https://data.example.com/vgi?a=<unsafe>");
        True(success.Contains("Authentication Successful") && success.Contains("prefers-color-scheme") && success.Contains("Cupola"), "OAuth success page uses the VGI visual treatment");
        True(success.Contains("Nearwater &lt;Production&gt;") && !success.Contains("Nearwater <Production>"), "OAuth success page escapes resource content");
        var failure = OAuthCallbackPage.Error("State mismatch", "Nearwater");
        True(failure.Contains("Authentication Failed") && failure.Contains("icon-circle-error"), "OAuth failure page uses the VGI visual treatment");
    }

    private static void TelemetryPrivacyTests()
    {
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

    private static void RibbonTests()
    {
        var xml = new VgiRibbon().GetCustomUI("Microsoft.Excel.Workbook");
        foreach (var expected in new[] { "label='Cupola'", "label='Cupola Data'", "label='Connections'", "label='Refresh Cupola tables'", "label='Refresh formulas'", "label='Diagnostics'" })
            True(xml.Contains(expected), "ribbon XML should contain " + expected);
        True(xml.Contains("getImage='GetCupolaImage'"), "primary ribbon control uses the Cupola image callback");
        True(!xml.Contains("imageMso='DatabaseInsert'"), "primary ribbon control no longer uses the generic database icon");
        True(!xml.Contains("Companion"), "ribbon must not expose the retired companion");
    }

    private static void BridgePolicyTests()
    {
        Equal("0.4.0", ProductInfo.Version, "native product version");
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
        Equal("Driver={Cupola for Excel};CupolaConnection={weather};", PowerQueryBridge.ConnectionString("weather", "Cupola for Excel"), "DSN-less Cupola ODBC contract");
        Equal("Driver={Driver}}Name};CupolaConnection={finance}}prod};", PowerQueryBridge.ConnectionString("finance}prod", "Driver}Name"), "ODBC brace escaping");
        var formula = PowerQueryBridge.Formula("SELECT \"amount\" FROM finance.main.ledger WHERE note = 'a'", "finance", "Haybarn VGI");
        True(formula.Contains("Odbc.Query(\"Driver={Haybarn VGI};CupolaConnection={finance};\""), "Power Query uses the configured ODBC driver and Cupola connection identity");
        True(formula.Contains("SELECT \"\"amount\"\" FROM finance.main.ledger"), "Power Query M string escaping");
        True(!formula.Contains("https://") && !formula.Contains("token") && !formula.Contains("secret"), "Power Query formula contains no endpoint or credential material");
        True(!PowerQueryBridge.IsDriverRegistered("Cupola test driver " + Guid.NewGuid().ToString("N")), "missing ODBC driver is detected before Excel creates a load sheet");
        Throws<InvalidOperationException>(() => PowerQueryBridge.Formula("COPY ledger TO 'C:\\ledger.csv'", "finance"), "Power Query export rejects mutating SQL");
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
        using var form = WebWorkbenchForm.Create(2);
        True(form is WebWorkbenchForm, "web assets should select the embedded Workbench");
        form.Show();
        for (var attempt = 0; attempt < 100 && WebWorkbenchForm.LastBridgeMethod != "ui.ready"; attempt++)
        {
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(100);
        }
        Equal("Ready", WebWorkbenchForm.LastStatus, "WebView2 navigation status");
        Equal("ui.ready", WebWorkbenchForm.LastBridgeMethod, "web/native bridge response handshake");
        form.Close();
    }

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
