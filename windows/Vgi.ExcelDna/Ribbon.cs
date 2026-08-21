using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace QueryFarm.Vgi.ExcelDna;

[ComVisible(true)]
public sealed class VgiRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='VgiTab' label='Cupola'>
        <group id='VgiDataGroup' label='Cupola Data'>
          <button id='VgiWorkbench' label='Cupola' size='large' getImage='GetCupolaImage' onAction='OpenWorkbench'/>
          <button id='VgiConnections' label='Connections' imageMso='DataConnections' onAction='OpenConnections'/>
          <button id='CupolaRefreshTables' label='Refresh Cupola tables' imageMso='RefreshAll' onAction='RefreshCupolaTables'/>
          <button id='VgiRefresh' label='Refresh formulas' imageMso='RefreshAll' onAction='RefreshFormulas'/>
        </group>
        <group id='VgiHelpGroup' label='Help'>
          <button id='VgiDiagnostics' label='Diagnostics' imageMso='Info' onAction='ShowDiagnostics'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public object GetCupolaImage(IRibbonControl control) => base.LoadImage("CupolaMark");

    public void OpenWorkbench(IRibbonControl control) => WorkbenchWindow.Show(0);
    public void OpenConnections(IRibbonControl control) => WorkbenchWindow.Show(3);
    public void RefreshCupolaTables(IRibbonControl control)
    {
        try { MessageBox.Show($"Refreshed {WorkbookBridge.RefreshAllSnapshots()} Cupola table(s).", ProductInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Information); }
        catch (Exception error) { ErrorLog.Write(error, "ribbon.refresh-snapshots"); MessageBox.Show(error.Message, ProductInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    public void RefreshFormulas(IRibbonControl control)
    {
        try { ((dynamic)ExcelDnaUtil.Application).CalculateFull(); }
        catch (Exception error) { ErrorLog.Write(error, "ribbon.refresh-formulas"); MessageBox.Show(error.Message, ProductInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    public void ShowDiagnostics(IRibbonControl control) =>
        MessageBox.Show(HaybarnClient.Diagnostics(), ProductInfo.Name + " diagnostics", MessageBoxButtons.OK, MessageBoxIcon.Information);
}

internal static class WorkbenchWindow
{
    private static Form? _form;

    public static void Show(int tab)
    {
        if (_form is null || _form.IsDisposed) _form = WebWorkbenchForm.Create(tab);
        if (_form is WebWorkbenchForm web) web.SelectTab(tab);
        else if (_form is NativeWorkbenchForm native) native.SelectTab(tab);
        if (!_form.Visible) _form.Show(new ExcelWindow(ExcelDnaUtil.WindowHandle));
        _form.BringToFront();
        _form.Activate();
    }

    private sealed class ExcelWindow : IWin32Window
    {
        public ExcelWindow(IntPtr handle) => Handle = handle;
        public IntPtr Handle { get; }
    }
}
