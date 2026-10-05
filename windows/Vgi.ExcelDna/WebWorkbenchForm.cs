using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Collections.Generic;
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
    private readonly Panel _opening = new() { Name = "openingPanel", Dock = DockStyle.Fill, BackColor = Color.FromArgb(247, 243, 234) };
    private readonly Label _openingText = new() { AutoSize = true, Text = "Opening Cupola…", ForeColor = Color.FromArgb(33, 26, 18), Font = new Font("Segoe UI", 12) };
    private readonly System.Windows.Forms.Timer _startupTimer = new() { Interval = 45_000 };
    private bool _startupFailed;
    private readonly int _initialTab;
    private readonly JObject? _resultsSnapshot;
    private bool _ready;
    private int _uiThreadId;
    internal static string LastStatus { get; private set; } = "Not started";
    internal static string LastBridgeMethod { get; private set; } = "";

    private WebWorkbenchForm(int tab)
    {
        _initialTab = tab;
        Text = ProductInfo.WindowTitle;
        CupolaWindowIcon.Apply(this);
        Width = 1060;
        Height = 760;
        MinimumSize = new Size(360, 480);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = _opening.BackColor;
        _web.DefaultBackgroundColor = BackColor;
        Controls.Add(_web);
        var mark = new PictureBox { Size = new Size(64, 64), SizeMode = PictureBoxSizeMode.Zoom, AccessibleName = "Cupola" };
        using (var stream = typeof(WebWorkbenchForm).Assembly.GetManifestResourceStream("QueryFarm.Vgi.ExcelDna.OpeningMark"))
            if (stream is not null) { using var image = System.Drawing.Image.FromStream(stream); mark.Image = new Bitmap(image); }
        _opening.Controls.Add(mark);
        _opening.Controls.Add(_openingText);
        _opening.Resize += (_, __) => {
            mark.Location = new Point(Math.Max(0, (_opening.Width - mark.Width) / 2), Math.Max(16, (_opening.Height - 110) / 2));
            _openingText.Location = new Point(Math.Max(0, (_opening.Width - _openingText.Width) / 2), mark.Bottom + 16);
        };
        Controls.Add(_opening);
        _opening.BringToFront();
        _startupTimer.Tick += (_, __) => StartupFailed(new TimeoutException("Cupola took too long to open. Close this window and try again, or open native Cupola below."));
        FormClosed += (_, __) => { _startupTimer.Dispose(); mark.Image?.Dispose(); };
        // Keep a native diagonal resize target outside the WebView child window.
        Controls.Add(new StatusStrip
        {
            Name = "windowResizeGrip",
            AccessibleName = "Window resize grip",
            Dock = DockStyle.Bottom,
            SizingGrip = true,
            RightToLeft = RightToLeft.No
        });
        Shown += async (_, __) => { _startupTimer.Start(); await InitializeWebView(); };
    }

    private WebWorkbenchForm(JObject snapshot) : this(0)
    {
        _resultsSnapshot = snapshot;
        _openingText.Text = "Opening results…";
        Text = (snapshot.Value<string>("title") ?? "Query") + " — Results — " + ProductInfo.WindowTitle;
        Width = 1200; Height = 850;
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
            var configOverride = Environment.GetEnvironmentVariable("VGI_EXCEL_CONFIG_HOME");
            var configRoot = string.IsNullOrWhiteSpace(configOverride) ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QueryFarm", "VgiExcel") : Path.GetFullPath(configOverride);
            var userData = Path.Combine(configRoot, "WebView2");
            Directory.CreateDirectory(userData);
            var environment = await CoreWebView2Environment.CreateAsync(null, userData);
            if (IsDisposed || _startupFailed) return;
            await _web.EnsureCoreWebView2Async(environment);
            if (IsDisposed || _startupFailed) return;
            _web.CoreWebView2.SetVirtualHostNameToFolderMapping(AppHost, AssetDirectory(), CoreWebView2HostResourceAccessKind.DenyCors);
            _web.CoreWebView2.Settings.AreDevToolsEnabled = Environment.GetEnvironmentVariable("VGI_EXCEL_WEBVIEW_DEVTOOLS") == "1";
            _web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _web.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _web.CoreWebView2.Settings.IsZoomControlEnabled = true;
            _web.CoreWebView2.WebMessageReceived += OnWebMessage;
            _web.CoreWebView2.NavigationCompleted += (_, args) => {
                if (!args.IsSuccess) StartupFailed(new InvalidOperationException("Cupola could not load its interface. Close this window and try again."));
            };
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
            _web.Source = new Uri(_resultsSnapshot is null ? $"https://{AppHost}/index.html?tab={TabName(_initialTab)}" : $"https://{AppHost}/results.html");
        }
        catch (Exception error)
        {
            StartupFailed(error);
        }
    }

    private void StartupFailed(Exception error)
    {
        if (IsDisposed || _startupFailed) return;
        _startupFailed = true;
        _startupTimer.Stop();
        LastStatus = "Failed: " + error.GetBaseException().Message;
        ErrorLog.Write(error);
        _web.Visible = false;
        if (_resultsSnapshot is null) ShowFallback(error);
        else _openingText.Text = "Could not open results. Close this window and try again.";
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
            object? result;
            if (method == "ui.rendered")
            {
                if (!_startupFailed) { _startupTimer.Stop(); _opening.Visible = false; LastStatus = "Ready"; }
                result = true;
            }
            else if (_resultsSnapshot is not null)
            {
                // Result viewers have no connection, query, credential, or workbook bridge.
                switch (method)
                {
                    case "results.ready": result = _resultsSnapshot; break;
                    case "results.maximize": WindowState = FormWindowState.Maximized; result = true; break;
                    case "results.close": Close(); return;
                    default: throw new InvalidOperationException("This action is unavailable in a results window.");
                }
            }
            else if (method == "results.open")
            {
                var snapshot = ValidateResultsSnapshot(parameters);
                var viewer = new WebWorkbenchForm(snapshot);
                viewer.Show(this);
                result = true;
            }
            else result = await WorkbenchBridge.Invoke(method, parameters, status => Reply(new JObject { ["id"] = id, ["progress"] = status }));
            LastBridgeMethod = method;
            Reply(new JObject { ["id"] = id, ["result"] = result is null ? JValue.CreateNull() : JToken.FromObject(result, WorkbenchBridge.Serializer) });
        }
        catch (Exception error)
        {
            ErrorLog.Write(error);
            Reply(new JObject { ["id"] = request?.Value<int?>("id") ?? 0, ["error"] = error.GetBaseException().Message });
        }
    }

    internal static JObject ValidateResultsSnapshot(JObject parameters)
    {
        var result = parameters["result"]?.ToObject<QueryResult>(WorkbenchBridge.Serializer)
            ?? throw new ArgumentException("Query results are required.");
        if (result.Columns is null || result.Rows is null || result.Rows.Length > 20_000 || result.RowCount < result.Rows.Length || result.Columns.Any(column => column is null) || result.Rows.Any(row => row is null || row.Length != result.Columns.Length))
            throw new ArgumentException("Query results are invalid.");
        return new JObject { ["title"] = parameters.Value<string>("title") ?? "Query", ["result"] = JToken.FromObject(result, WorkbenchBridge.Serializer) };
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
        _web.Dispose();
        _opening.Dispose();
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
    private static readonly object queryGate = new();
    private static readonly Dictionary<string, CancellationTokenSource> editorQueries = new();

    internal static readonly JsonSerializer Serializer = JsonSerializer.Create(new JsonSerializerSettings
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        NullValueHandling = NullValueHandling.Include
    });

    public static async Task<object?> Invoke(string method, JObject parameters, Action<string>? progress = null)
    {
        switch (method)
        {
            case "app.info": return new { ProductInfo.Name, ProductInfo.Version, ProductInfo.Build };
            case "app.diagnostics": return HaybarnClient.Diagnostics();
            case "updates.status": return parameters.Value<bool>("autoCheck") ? await UpdateService.AutoCheck() : UpdateService.Status();
            case "updates.preference": return UpdateService.Preference(parameters.Value<bool>("daily"));
            case "updates.check": return await UpdateService.Check();
            case "updates.download": return await UpdateService.Download(progress);
            case "updates.install": return await UpdateService.Install();
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
                ConnectionStore.Save(connection, parameters.Value<bool?>("makeDefault") ?? false, parameters.Value<string>("originalName") ?? "");
                return Connections();
            }
            case "connections.workspace":
                ConnectionStore.SetWorkspace(parameters["members"]?.ToObject<string[]>(Serializer) ?? Array.Empty<string>());
                return Connections();
            case "connections.use":
                ConnectionStore.SetDefault(parameters.Value<string>("name") ?? "");
                return Connections();
            case "connections.remove":
            {
                var name = parameters.Value<string>("name") ?? "";
                var existing = ConnectionStore.List().FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                ConnectionStore.Remove(name);
                if (existing is not null && !existing.IsProfile) OAuthClient.SignOut(existing);
                return Connections();
            }
            case "connections.catalogs":
                return await ConnectionProbe.Catalogs(parameters.Value<string>("location") ?? "", progress);
            case "connections.test":
            {
                var connection = RequiredConnection(parameters);
                await ConnectionProbe.Test(connection);
                return new { authentication = connection.Authentication };
            }
            case "connections.signIn":
            {
                var connection = RequiredConnection(parameters);
                ConnectionStore.Save(connection);
                foreach (var member in ConnectionStore.ResolveAttachments(connection))
                {
                    member.Authentication = "oauth";
                    ConnectionStore.Save(member, false);
                    await OAuthClient.SignInAsync(member);
                }
                return Connections();
            }
            case "connections.signOut":
            {
                var connection = RequiredConnection(parameters);
                foreach (var member in ConnectionStore.ResolveAttachments(connection)) OAuthClient.SignOut(member);
                return Connections();
            }
            case "query.cancel":
            {
                var id = parameters.Value<string>("queryId") ?? "";
                lock (queryGate)
                {
                    if (!editorQueries.TryGetValue(id, out var source)) return false;
                    source.Cancel();
                    return true;
                }
            }
            case "query.agent":
            case "query.editor":
            {
                var id = parameters.Value<string>("queryId") ?? "";
                if (!Guid.TryParse(id, out _)) throw new ArgumentException("A query identifier is required.");
                var sql = parameters.Value<string>("sql") ?? "";
                if (method == "query.agent" || parameters.Value<bool?>("agent") == true) AgentSqlPolicy.AssertReadOnly(sql);
                using var source = new CancellationTokenSource();
                lock (queryGate)
                {
                    if (editorQueries.ContainsKey(id)) throw new ArgumentException("Query is already running.");
                    editorQueries.Add(id, source);
                }
                try { return await Task.Run(() => new HaybarnClient().QueryResult(sql, parameters.Value<string>("connection"), Math.Max(1, Math.Min(20_000, parameters.Value<int?>("maxRows") ?? 10_000)), source.Token)); }
                catch (OperationCanceledException) when (source.IsCancellationRequested) { return null; }
                finally { lock (queryGate) editorQueries.Remove(id); }
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
                return WorkbookBridge.InsertAtActiveCell(result, parameters.Value<string>("tableName") ?? "VGI_Result");
            }
            case "excel.insertQuery":
            {
                var sql = parameters.Value<string>("sql") ?? "";
                var connection = parameters.Value<string>("connection");
                AgentSqlPolicy.AssertReadOnly(sql);
                var result = await Task.Run(() => new HaybarnClient().QueryResult(sql, connection, WorkbookBridge.MaximumWorksheetDataRows + 1));
                if (result.Truncated || result.RowCount > WorkbookBridge.MaximumWorksheetDataRows)
                    throw new InvalidOperationException($"The query returned {result.RowCount:N0} rows. Excel tables can contain at most {WorkbookBridge.MaximumWorksheetDataRows:N0} data rows on a worksheet.");
                return WorkbookBridge.InsertAtActiveCell(result, parameters.Value<string>("tableName") ?? "VGI_Result");
            }
            case "excel.createPowerQuery":
                return PowerQueryBridge.Create(
                    parameters.Value<string>("sql") ?? "",
                    parameters.Value<string>("connection") ?? "",
                    parameters.Value<string>("name"),
                    parameters.Value<bool?>("loadToWorksheet") ?? true,
                    parameters.Value<string>("sheetName"),
                    parameters.Value<string>("tableName"));
            case "excel.activateTable": return WorkbookBridge.ActivateTable(parameters.Value<string>("tableName") ?? "");
            case "excel.snapshots": return WorkbookBridge.ManagedSnapshots();
            case "excel.refreshSnapshot": return WorkbookBridge.RefreshSnapshot(parameters.Value<string>("tableName") ?? "");
            case "excel.migrateSnapshot": return WorkbookBridge.CreateRefreshableCopy(parameters.Value<string>("tableName") ?? "");
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
        var workspace = ConnectionStore.WorkspaceName();
        var saved = ConnectionStore.List();
        return saved.Select(connection =>
        {
            IReadOnlyList<VgiConnection> members;
            try { members = ConnectionStore.ResolveAttachments(connection, saved); }
            catch { members = Array.Empty<VgiConnection>(); }
            return (object)new
            {
                connection.Name, Catalog = members.FirstOrDefault()?.Catalog ?? connection.Catalog,
                Catalogs = members.Select(member => member.Catalog).ToArray(), connection.Members,
                connection.Location, connection.Authentication, connection.AttachOptions, connection.IsWorkspaceProfile,
                IsWorkspaceSelected = string.Equals(connection.Name, workspace, StringComparison.OrdinalIgnoreCase),
                IsDefault = string.Equals(connection.Name, preferred, StringComparison.OrdinalIgnoreCase),
                IsSignedIn = members.Any(member => member.Authentication == "oauth") && members.All(member => member.Authentication != "oauth" || OAuthClient.IsSignedIn(member))
            };
        }).ToArray();
    }

    private static VgiConnection RequiredConnection(JObject parameters) =>
        parameters["connection"]?.ToObject<VgiConnection>(Serializer) ?? throw new ArgumentException("A connection is required.");

}
