using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class NativeWorkbenchForm : Form
{
    private readonly HaybarnClient _client = new();
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };
    private readonly ComboBox _queryConnection = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly TextBox _sql = new() { Multiline = true, AcceptsReturn = true, AcceptsTab = true, ScrollBars = ScrollBars.Both, Dock = DockStyle.Fill, Font = new Font("Consolas", 10), Text = "SELECT current_catalog(), current_schema();" };
    private readonly DataGridView _results = Grid();
    private readonly Label _status = new() { AutoSize = true, Text = "Ready", Padding = new Padding(8, 7, 0, 0) };
    private QueryResult? _lastResult;

    private readonly ComboBox _catalogConnection = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly DataGridView _catalog = Grid();
    private readonly ComboBox _agentConnection = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly TextBox _agentKey = new() { Width = 260, UseSystemPasswordChar = true };
    private readonly TextBox _agentPrompt = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
    private readonly RichTextBox _agentAnswer = new() { ReadOnly = true, BorderStyle = BorderStyle.None, ScrollBars = RichTextBoxScrollBars.Vertical, Dock = DockStyle.Fill, BackColor = SystemColors.Window, Font = new Font("Segoe UI", 10), DetectUrls = true };
    private readonly DataGridView _agentResult = Grid();

    private readonly ListBox _connections = new() { Dock = DockStyle.Fill };
    private readonly TextBox _name = new() { Width = 210 };
    private readonly TextBox _catalogName = new() { Width = 210 };
    private readonly TextBox _location = new() { Width = 420 };
    private string _connectionAuthentication = "anonymous";
    private readonly StatusStrip _productStatus = new() { SizingGrip = false };

    public NativeWorkbenchForm()
    {
        Text = ProductInfo.Name;
        Width = 960;
        Height = 680;
        MinimumSize = new Size(720, 500);
        StartPosition = FormStartPosition.CenterScreen;
        _tabs.TabPages.Add(BuildQueryTab());
        _tabs.TabPages.Add(BuildCatalogTab());
        _tabs.TabPages.Add(BuildAgentTab());
        _tabs.TabPages.Add(BuildConnectionsTab());
        Controls.Add(_tabs);
        _productStatus.Items.Add(new ToolStripStatusLabel
        {
            Spring = true,
            TextAlign = ContentAlignment.MiddleRight,
            Text = $"{ProductInfo.Name} v{ProductInfo.Version} · build {ProductInfo.Build} · Query.Farm"
        });
        Controls.Add(_productStatus);
        _agentKey.Text = AgentCredentialStore.Load() ?? "";
        RefreshConnections();
    }

    public void SelectTab(int index) => _tabs.SelectedIndex = Math.Max(0, Math.Min(index, _tabs.TabPages.Count - 1));

    private TabPage BuildQueryTab()
    {
        var page = new TabPage("SQL");
        var run = Button("Run", async (_, __) => await RunQuery());
        var insert = Button("Insert as table", (_, __) => InsertResult());
        var toolbar = Flow(new Label { Text = "Connection", AutoSize = true, Padding = new Padding(0, 7, 0, 0) }, _queryConnection, run, insert, _status);
        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 210 };
        split.Panel1.Controls.Add(_sql);
        split.Panel2.Controls.Add(_results);
        page.Controls.Add(split);
        page.Controls.Add(toolbar);
        toolbar.Dock = DockStyle.Top;
        return page;
    }

    private TabPage BuildCatalogTab()
    {
        var page = new TabPage("Catalog");
        var refresh = Button("Explore", async (_, __) => await ExploreCatalog());
        var use = Button("Open in SQL", (_, __) => OpenSelectedCatalogObject());
        var toolbar = Flow(new Label { Text = "Connection", AutoSize = true, Padding = new Padding(0, 7, 0, 0) }, _catalogConnection, refresh, use);
        _catalog.CellDoubleClick += (_, __) => OpenSelectedCatalogObject();
        page.Controls.Add(_catalog);
        page.Controls.Add(toolbar);
        toolbar.Dock = DockStyle.Top;
        return page;
    }

    private TabPage BuildConnectionsTab()
    {
        var page = new TabPage("Connections");
        var form = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12), ColumnCount = 2 };
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        AddRow(form, "Name", _name);
        AddRow(form, "VGI catalog", _catalogName);
        AddRow(form, "HTTPS endpoint", _location);
        var save = Button("Save and use", (_, __) => SaveConnection());
        var remove = Button("Remove", (_, __) => RemoveConnection());
        var test = Button("Test", async (_, __) => await TestConnection());
        var signOut = Button("Sign out", (_, __) => SignOut());
        form.Controls.Add(Flow(save, test, signOut, remove), 1, 3);
        var hint = new Label { Dock = DockStyle.Top, Height = 42, Padding = new Padding(12, 5, 12, 5), Text = "Only HTTPS VGI services are supported. If authentication is required, Cupola opens your browser automatically." };
        _connections.SelectedIndexChanged += (_, __) => LoadSelectedConnection();
        page.Controls.Add(_connections);
        page.Controls.Add(hint);
        page.Controls.Add(form);
        hint.Dock = DockStyle.Top;
        return page;
    }

    private TabPage BuildAgentTab()
    {
        var page = new TabPage("Agent");
        var ask = Button("Ask", async (_, __) => await AskAgent());
        var insert = Button("Insert preview", (_, __) => InsertResult());
        var saveKey = Button("Save key", (_, __) => SaveAgentKey());
        var forgetKey = Button("Forget", (_, __) => ForgetAgentKey());
        var toolbar = Flow(new Label { Text = "Connection", AutoSize = true, Padding = new Padding(0, 7, 0, 0) }, _agentConnection,
            new Label { Text = "Anthropic key", AutoSize = true, Padding = new Padding(8, 7, 0, 0) }, _agentKey, saveKey, forgetKey, ask, insert);
        var outer = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 150 };
        outer.Panel1.Controls.Add(_agentPrompt);
        var lower = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 120 };
        lower.Panel1.Controls.Add(_agentAnswer);
        lower.Panel2.Controls.Add(_agentResult);
        outer.Panel2.Controls.Add(lower);
        page.Controls.Add(outer);
        page.Controls.Add(toolbar);
        toolbar.Dock = DockStyle.Top;
        return page;
    }

    private async Task AskAgent()
    {
        var connection = SelectedName(_agentConnection);
        if (connection is null) { SetStatus("Add a connection first."); return; }
        await Busy(async () =>
        {
            var answer = await Task.Run(() => new AgentClient(_client).RunAsync(_agentKey.Text, _agentPrompt.Text, connection));
            MarkdownRichText.Set(_agentAnswer, answer.Text);
            _lastResult = answer.StagedResult;
            if (_lastResult is not null) ShowResult(_agentResult, _lastResult);
            SetStatus(_lastResult is null ? "Agent finished." : "Agent preview is ready for review.");
        });
    }

    private void SaveAgentKey()
    {
        try { AgentCredentialStore.Save(_agentKey.Text); SetStatus("Anthropic API key saved in Windows Credential Manager."); }
        catch (Exception error) { ErrorLog.Write(error); SetStatus(error.Message); }
    }

    private void ForgetAgentKey()
    {
        try { AgentCredentialStore.Delete(); _agentKey.Clear(); SetStatus("Saved Anthropic API key removed."); }
        catch (Exception error) { ErrorLog.Write(error); SetStatus(error.Message); }
    }

    private async Task RunQuery()
    {
        var connection = SelectedName(_queryConnection);
        if (connection is null) { SetStatus("Add a connection first."); return; }
        await Busy(async () =>
        {
            _lastResult = await Task.Run(() => _client.QueryResult(_sql.Text, connection, 10_000));
            ShowResult(_results, _lastResult);
            SetStatus($"{_lastResult.RowCount:N0} rows · {_lastResult.ElapsedMs:N0} ms");
        });
    }

    private async Task ExploreCatalog()
    {
        var connection = SelectedName(_catalogConnection);
        if (connection is null) { MessageBox.Show(this, "Add a connection first.", ProductInfo.Name); return; }
        const string sql = @"SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, table_type AS kind, CASE WHEN table_type = 'VIEW' THEN 'view' ELSE 'table' END AS object_type
FROM information_schema.tables WHERE table_schema NOT IN ('information_schema', 'pg_catalog')
UNION ALL
SELECT database_name, schema_name, function_name, function_type, CASE WHEN function_type IN ('macro', 'table_macro') THEN 'macro' ELSE 'function' END FROM duckdb_functions()
WHERE database_name NOT IN ('system', 'temp')
ORDER BY 1, 2, 3";
        await Busy(async () => ShowResult(_catalog, await Task.Run(() => _client.QueryResult(sql, connection, 20_000))));
    }

    private void OpenSelectedCatalogObject()
    {
        if (_catalog.CurrentRow is null || _catalog.CurrentRow.Cells.Count < 4) return;
        var catalog = Convert.ToString(_catalog.CurrentRow.Cells[0].Value) ?? "";
        var schema = Convert.ToString(_catalog.CurrentRow.Cells[1].Value) ?? "";
        var name = Convert.ToString(_catalog.CurrentRow.Cells[2].Value) ?? "";
        var kind = Convert.ToString(_catalog.CurrentRow.Cells[3].Value) ?? "";
        var objectType = _catalog.CurrentRow.Cells.Count > 4 ? Convert.ToString(_catalog.CurrentRow.Cells[4].Value) ?? "" : "";
        _sql.Text = objectType == "table" || objectType == "view"
            ? $"SELECT * FROM {Quote(catalog)}.{Quote(schema)}.{Quote(name)} LIMIT 100;"
            : $"-- {catalog}.{schema}.{name} ({kind})\r\nSELECT f.function_type, f.description, f.comment, f.return_type, to_json(f.examples) AS examples, to_json(f.tags) AS tags, a.arg_position, a.arg_name, a.arg_type, a.is_named, a.is_positional, a.arg_default, a.arg_choices, a.arg_range, a.arg_pattern, a.arg_description FROM duckdb_functions() f LEFT JOIN vgi_function_arguments() a ON a.catalog_name = f.database_name AND a.schema_name = f.schema_name AND a.function_name = f.function_name WHERE f.database_name = '{Literal(catalog)}' AND f.schema_name = '{Literal(schema)}' AND f.function_name = '{Literal(name)}' ORDER BY a.field_index;";
        _tabs.SelectedIndex = 0;
    }

    private void SaveConnection()
    {
        try
        {
            ConnectionStore.Save(CurrentConnection());
            RefreshConnections();
            MessageBox.Show(this, "The HTTPS VGI connection was saved.", ProductInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception error) { ErrorLog.Write(error, "connection.save"); MessageBox.Show(this, error.Message, "VGI connection", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private async Task TestConnection()
    {
        try
        {
            var candidate = CurrentConnection();
            ConnectionStore.Validate(candidate);
            ConnectionStore.Save(candidate);
            var message = "";
            var succeeded = await Busy(async () =>
            {
                var result = await Task.Run(() => _client.QueryResult("SELECT current_catalog(), current_schema();", candidate.Name, 1));
                message = result.Rows.Length == 1 ? "Connection successful." : "The connection returned no result.";
            });
            if (succeeded) MessageBox.Show(this, message, "VGI connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            RefreshConnections();
        }
        catch (Exception error) { ErrorLog.Write(error, "connection.test"); MessageBox.Show(this, error.Message, "VGI connection", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private void SignOut()
    {
        try
        {
            var connection = CurrentConnection();
            OAuthClient.SignOut(connection);
            MessageBox.Show(this, "The saved OAuth session was removed.", "VGI connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception error) { ErrorLog.Write(error, "connection.sign-out"); MessageBox.Show(this, error.Message, "VGI sign-out", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private void RemoveConnection()
    {
        var name = Convert.ToString(_connections.SelectedItem);
        if (string.IsNullOrWhiteSpace(name)) return;
        ConnectionStore.Remove(name);
        RefreshConnections();
    }

    private void LoadSelectedConnection()
    {
        var selected = Convert.ToString(_connections.SelectedItem);
        var value = ConnectionStore.List().FirstOrDefault(item => item.Name == selected);
        if (value is null) return;
        _name.Text = value.Name;
        _catalogName.Text = value.Catalog;
        _location.Text = value.Location;
        _connectionAuthentication = string.IsNullOrWhiteSpace(value.Authentication) ? "anonymous" : value.Authentication;
    }

    private void RefreshConnections()
    {
        var values = ConnectionStore.List();
        var preferred = ConnectionStore.DefaultName();
        _connections.Items.Clear();
        _queryConnection.Items.Clear();
        _catalogConnection.Items.Clear();
        _agentConnection.Items.Clear();
        foreach (var item in values)
        {
            _connections.Items.Add(item.Name);
            _queryConnection.Items.Add(item.Name);
            _catalogConnection.Items.Add(item.Name);
            _agentConnection.Items.Add(item.Name);
        }
        Select(_queryConnection, preferred);
        Select(_catalogConnection, preferred);
        Select(_agentConnection, preferred);
        if (_connections.Items.Count > 0) _connections.SelectedIndex = Math.Max(0, _connections.Items.IndexOf(preferred));
    }

    private async Task<bool> Busy(Func<Task> action)
    {
        UseWaitCursor = true;
        SetStatus("Working…");
        try { await action(); return true; }
        catch (Exception error) { ErrorLog.Write(error); SetStatus(error.Message); MessageBox.Show(this, error.Message, ProductInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
        finally { UseWaitCursor = false; }
    }

    private void InsertResult()
    {
        if (_lastResult is null) return;
        try { WorkbookBridge.InsertAtActiveCell(_lastResult, "VGI_Result"); }
        catch (Exception error) { ErrorLog.Write(error, "excel.insert"); MessageBox.Show(this, error.Message, "Insert VGI result", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private static void ShowResult(DataGridView grid, QueryResult result)
    {
        grid.Rows.Clear(); grid.Columns.Clear();
        foreach (var column in result.Columns) grid.Columns.Add(column.Name, $"{column.Name}  [{column.Type}]");
        foreach (var row in result.Rows) grid.Rows.Add(row.Cast<object>().ToArray());
    }

    private static DataGridView Grid() => new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
    private static FlowLayoutPanel Flow(params Control[] controls) { var value = new FlowLayoutPanel { AutoSize = true, Padding = new Padding(8), WrapContents = false }; value.Controls.AddRange(controls); return value; }
    private static Button Button(string text, EventHandler click) { var value = new Button { Text = text, AutoSize = true }; value.Click += click; return value; }
    private static void AddRow(TableLayoutPanel form, string label, Control control) { var row = form.RowCount++; form.Controls.Add(new Label { Text = label, AutoSize = true, Padding = new Padding(0, 7, 0, 0) }, 0, row); form.Controls.Add(control, 1, row); }
    private static string? SelectedName(ComboBox box) => Convert.ToString(box.SelectedItem);
    private static void Select(ComboBox box, string? name) { if (box.Items.Count == 0) return; var index = string.IsNullOrWhiteSpace(name) ? -1 : box.Items.IndexOf(name); box.SelectedIndex = index >= 0 ? index : 0; }
    private static string Quote(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    private static string Literal(string value) => value.Replace("'", "''");
    private VgiConnection CurrentConnection() => new()
    {
        Name = _name.Text.Trim(), Catalog = _catalogName.Text.Trim(), Location = _location.Text.Trim(),
        Authentication = _connectionAuthentication,
        AttachOptions = ConnectionStore.List().FirstOrDefault(item => string.Equals(item.Name, _name.Text.Trim(), StringComparison.OrdinalIgnoreCase))?.AttachOptions
            ?? new System.Collections.Generic.Dictionary<string, object?>()
    };
    private void SetStatus(string value) => _status.Text = value;
}
