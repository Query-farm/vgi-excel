using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class WebWorkbenchForm : Form
{
    private const string AppHost = "vgi-excel.local";
    private readonly WebView2 _web = new() { Dock = DockStyle.Fill };
    private readonly int _initialTab;
    private bool _ready;
    private int _uiThreadId;
    internal static string LastStatus { get; private set; } = "Not started";
    internal static string LastBridgeMethod { get; private set; } = "";

    private WebWorkbenchForm(int tab)
    {
        _initialTab = tab;
        Text = ProductInfo.Name;
        Width = 1060;
        Height = 760;
        MinimumSize = new Size(360, 480);
        StartPosition = FormStartPosition.CenterParent;
        Controls.Add(_web);
        Shown += async (_, __) => await InitializeWebView();
    }

    public static Form Create(int tab)
    {
        var assets = AssetDirectory();
        if (Environment.GetEnvironmentVariable("VGI_EXCEL_NATIVE_WORKBENCH") == "1" || !Directory.Exists(assets))
        {
            LastStatus = Environment.GetEnvironmentVariable("VGI_EXCEL_NATIVE_WORKBENCH") == "1"
                ? "Native fallback requested"
                : "Native fallback: web assets missing at " + assets;
            var native = new NativeWorkbenchForm();
            native.SelectTab(tab);
            return native;
        }
        return new WebWorkbenchForm(tab);
    }

    public async void SelectTab(int tab)
    {
        if (!_ready) return;
        try { await _web.CoreWebView2.ExecuteScriptAsync($"window.vgiSelectTab && window.vgiSelectTab({JsonConvert.SerializeObject(TabName(tab))})"); }
        catch (Exception error) { ErrorLog.Write(error); }
    }

    private async Task InitializeWebView()
    {
        try
        {
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            LastStatus = "Initializing";
            var userData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QueryFarm", "VgiExcel", "WebView2");
            Directory.CreateDirectory(userData);
            var environment = await CoreWebView2Environment.CreateAsync(null, userData);
            await _web.EnsureCoreWebView2Async(environment);
            _web.CoreWebView2.SetVirtualHostNameToFolderMapping(AppHost, AssetDirectory(), CoreWebView2HostResourceAccessKind.DenyCors);
            _web.CoreWebView2.Settings.AreDevToolsEnabled = Environment.GetEnvironmentVariable("VGI_EXCEL_WEBVIEW_DEVTOOLS") == "1";
            _web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _web.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _web.CoreWebView2.Settings.IsZoomControlEnabled = true;
            _web.CoreWebView2.WebMessageReceived += OnWebMessage;
            _web.CoreWebView2.NavigationCompleted += (_, args) => LastStatus = args.IsSuccess ? "Ready" : $"Navigation failed: {args.WebErrorStatus}";
            _web.CoreWebView2.NavigationStarting += (_, args) =>
            {
                if (!args.Uri.StartsWith($"https://{AppHost}/", StringComparison.OrdinalIgnoreCase)) args.Cancel = true;
            };
            _web.CoreWebView2.NewWindowRequested += (_, args) =>
            {
                args.Handled = true;
                if (Uri.TryCreate(args.Uri, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps)
                    Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
            };
            _ready = true;
            _web.Source = new Uri($"https://{AppHost}/index.html?tab={TabName(_initialTab)}");
        }
        catch (Exception error)
        {
            LastStatus = "Failed: " + error.GetBaseException().Message;
            ErrorLog.Write(error);
            ShowFallback(error);
        }
    }

    private async void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        JObject? request = null;
        try
        {
            if (!args.Source.StartsWith($"https://{AppHost}/", StringComparison.OrdinalIgnoreCase)) return;
            request = JObject.Parse(args.WebMessageAsJson);
            var id = request.Value<int>("id");
            var method = request.Value<string>("method") ?? "";
            var parameters = request["params"] as JObject ?? new JObject();
            var result = await WorkbenchBridge.Invoke(method, parameters);
            LastBridgeMethod = method;
            Reply(new JObject { ["id"] = id, ["result"] = result is null ? JValue.CreateNull() : JToken.FromObject(result, WorkbenchBridge.Serializer) });
        }
        catch (Exception error)
        {
            ErrorLog.Write(error);
            Reply(new JObject { ["id"] = request?.Value<int?>("id") ?? 0, ["error"] = error.GetBaseException().Message });
        }
    }

    private void Reply(JObject value)
    {
        if (IsDisposed || _web.IsDisposed) return;
        if (_uiThreadId != 0 && Thread.CurrentThread.ManagedThreadId != _uiThreadId)
        {
            try { BeginInvoke(new Action(() => Reply(value))); }
            catch (Exception error) { ErrorLog.Write(error); }
            return;
        }
        try
        {
            if (_ready && _web.CoreWebView2 is not null)
            {
                var json = value.ToString(Formatting.None);
                _web.CoreWebView2.PostWebMessageAsJson(json);

                // Keep a direct JS callback as a compatibility path for WebView2
                // runtimes that expose a posted JSON message as a string. Duplicate
                // responses are harmless because the browser removes completed IDs.
                var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
                _ = _web.CoreWebView2.ExecuteScriptAsync(
                    $"window.vgiReceiveHostResponse && window.vgiReceiveHostResponse(JSON.parse(new TextDecoder().decode(Uint8Array.from(atob('{encoded}'), c => c.charCodeAt(0)))));"
                );
            }
        }
        catch (Exception error) { ErrorLog.Write(error); }
    }

    private void ShowFallback(Exception error)
    {
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(() => ShowFallback(error))); }
            catch (InvalidOperationException) { }
            return;
        }
        Controls.Clear();
        var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(24), WrapContents = false };
        panel.Controls.Add(new Label { AutoSize = true, Font = new Font("Segoe UI", 14, FontStyle.Bold), Text = "The modern Cupola for Excel experience could not start." });
        panel.Controls.Add(new Label { AutoSize = true, MaximumSize = new Size(760, 0), Text = error.GetBaseException().Message });
        var open = new Button { AutoSize = true, Text = "Open native Cupola" };
        open.Click += (_, __) => { var native = new NativeWorkbenchForm(); native.SelectTab(_initialTab); native.Show(); Close(); };
        panel.Controls.Add(open);
        Controls.Add(panel);
    }

    private static string AssetDirectory()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_EXCEL_WEB_ASSETS_PATH");
        return string.IsNullOrWhiteSpace(configured)
            ? Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? AppDomain.CurrentDomain.BaseDirectory, "web")
            : Path.GetFullPath(configured);
    }
    private static string TabName(int tab) => tab switch { 1 => "catalog", 2 => "agent", 3 => "connections", _ => "sql" };
}

internal static class WorkbenchBridge
{
    internal static readonly JsonSerializer Serializer = JsonSerializer.Create(new JsonSerializerSettings
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        NullValueHandling = NullValueHandling.Include
    });

    public static async Task<object?> Invoke(string method, JObject parameters)
    {
        switch (method)
        {
            case "app.info": return new { ProductInfo.Name, ProductInfo.Version, ProductInfo.Build };
            case "app.diagnostics": return HaybarnClient.Diagnostics();
            case "ui.ready": return true;
            case "agent.key.load": return AgentCredentialStore.Load();
            case "agent.key.save":
                AgentCredentialStore.Save(parameters.Value<string>("key") ?? "");
                return true;
            case "agent.key.delete":
                AgentCredentialStore.Delete();
                return true;
            case "agent.trace":
                AgentTraceLog.Write(parameters["event"] as JObject ?? new JObject { ["event"] = "invalid_trace" });
                return true;
            case "clipboard.write":
                Clipboard.SetText(parameters.Value<string>("value") ?? "");
                return true;
            case "connections.list": return Connections();
            case "connections.save":
            {
                var connection = RequiredConnection(parameters);
                ConnectionStore.Save(connection, parameters.Value<bool?>("makeDefault") ?? false);
                return Connections();
            }
            case "connections.use":
                ConnectionStore.SetDefault(parameters.Value<string>("name") ?? "");
                return Connections();
            case "connections.remove":
            {
                var name = parameters.Value<string>("name") ?? "";
                var existing = ConnectionStore.List().FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing is not null) OAuthClient.SignOut(existing);
                ConnectionStore.Remove(name);
                return Connections();
            }
            case "connections.test":
            {
                var connection = RequiredConnection(parameters);
                ConnectionStore.Save(connection);
                return await Task.Run(() => new HaybarnClient().QueryResult("SELECT current_catalog(), current_schema();", connection.Name, 1));
            }
            case "connections.signIn":
            {
                var connection = RequiredConnection(parameters);
                connection.Authentication = "oauth";
                ConnectionStore.Save(connection);
                await OAuthClient.SignInAsync(connection);
                return Connections();
            }
            case "connections.signOut":
            {
                var connection = RequiredConnection(parameters);
                OAuthClient.SignOut(connection);
                return Connections();
            }
            case "query.run":
            {
                var sql = parameters.Value<string>("sql") ?? "";
                if (parameters.Value<bool?>("agent") == true) AgentSqlPolicy.AssertReadOnly(sql);
                var connection = parameters.Value<string>("connection");
                var maxRows = Math.Max(1, Math.Min(20_000, parameters.Value<int?>("maxRows") ?? 10_000));
                return await Task.Run(() => new HaybarnClient().QueryResult(sql, connection, maxRows));
            }
            case "excel.insert":
            {
                var result = parameters["result"]?.ToObject<QueryResult>(Serializer) ?? throw new ArgumentException("A query result is required.");
                return WorkbookBridge.InsertAtActiveCell(result, parameters.Value<string>("tableName") ?? "VGI_Result", parameters.Value<string>("sql"), parameters.Value<string>("connection"));
            }
            case "excel.insertQuery":
            {
                var sql = parameters.Value<string>("sql") ?? "";
                var connection = parameters.Value<string>("connection");
                AgentSqlPolicy.AssertReadOnly(sql);
                var result = await Task.Run(() => new HaybarnClient().QueryResult(sql, connection, WorkbookBridge.MaximumWorksheetDataRows + 1));
                if (result.Truncated || result.RowCount > WorkbookBridge.MaximumWorksheetDataRows)
                    throw new InvalidOperationException($"The query returned {result.RowCount:N0} rows. Excel tables can contain at most {WorkbookBridge.MaximumWorksheetDataRows:N0} data rows on a worksheet.");
                return WorkbookBridge.InsertAtActiveCell(result, parameters.Value<string>("tableName") ?? "VGI_Result", sql, connection);
            }
            case "excel.createPowerQuery":
                return PowerQueryBridge.Create(
                    parameters.Value<string>("sql") ?? "",
                    parameters.Value<string>("connection") ?? "",
                    parameters.Value<string>("name"),
                    parameters.Value<bool?>("loadToWorksheet") ?? true);
            case "excel.activateTable": return WorkbookBridge.ActivateTable(parameters.Value<string>("tableName") ?? "");
            case "excel.snapshots": return WorkbookBridge.ManagedSnapshots();
            case "excel.refreshSnapshot": return WorkbookBridge.RefreshSnapshot(parameters.Value<string>("tableName") ?? "");
            case "excel.forgetSnapshot": return WorkbookBridge.ForgetSnapshot(parameters.Value<string>("tableName") ?? "");
            case "excel.workbookOverview": return WorkbookBridge.Overview();
            case "excel.readRange": return WorkbookBridge.ReadRange(parameters.Value<string>("sheet") ?? "", parameters.Value<string>("address") ?? "");
            case "excel.listFormulas": return WorkbookBridge.ListFormulas(parameters.Value<string>("sheet"), parameters.Value<int?>("limit") ?? 200);
            case "excel.writeResult":
            {
                var result = parameters["result"]?.ToObject<QueryResult>(Serializer) ?? throw new ArgumentException("A query result is required.");
                return WorkbookBridge.WriteResult(parameters.Value<string>("mode") ?? "", result, parameters.Value<string>("sheetName"), parameters.Value<string>("tableName") ?? "VGI_Result");
            }
            default: throw new ArgumentException($"Unknown Workbench operation: {method}");
        }
    }

    private static object[] Connections()
    {
        var preferred = ConnectionStore.DefaultName();
        return ConnectionStore.List().Select(connection => (object)new
        {
            connection.Name, connection.Catalog, connection.Location, connection.Authentication, connection.AttachOptions,
            IsDefault = string.Equals(connection.Name, preferred, StringComparison.OrdinalIgnoreCase),
            IsSignedIn = connection.Authentication == "oauth" && OAuthClient.IsSignedIn(connection)
        }).ToArray();
    }

    private static VgiConnection RequiredConnection(JObject parameters) =>
        parameters["connection"]?.ToObject<VgiConnection>(Serializer) ?? throw new ArgumentException("A connection is required.");

}
