using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: ComVisible(false)]
namespace QueryFarm.Cupola.ExcelLoader;

// Office's shared COM add-in startup interface. No VSTO runtime is required.
[ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IDTExtensibility2
{
    [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInstance, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(2)] void OnDisconnection(int removeMode, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(3)] void OnAddInsUpdate([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(4)] void OnStartupComplete([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(5)] void OnBeginShutdown([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
}

[ComVisible(true), Guid("59B35EE7-3344-4F74-94C6-030FE4AA236C"), ProgId("QueryFarm.Cupola.ExcelLoader"), ClassInterface(ClassInterfaceType.None)]
public sealed class Connect : IDTExtensibility2
{
    private object? application;
    private bool loaded;
    public void OnConnection(object application, int connectMode, object addInInstance, ref Array custom)
    {
        this.application = application;
        if (connectMode != 1) Load(); // ext_cm_Startup waits until Excel finishes starting.
    }
    public void OnStartupComplete(ref Array custom) => Load();
    public void OnDisconnection(int removeMode, ref Array custom) { application = null; }
    public void OnBeginShutdown(ref Array custom) { application = null; }
    public void OnAddInsUpdate(ref Array custom) { }

    private void Load()
    {
        if (loaded || application == null) return;
        // Only load the XLL shipped beside this registered, machine-installed assembly.
        var xll = Path.Combine(Path.GetDirectoryName(typeof(Connect).Assembly.Location)!, "Vgi.ExcelDna64-packed.xll");
        if (!File.Exists(xll)) throw new COMException("Cupola installation is incomplete. Repair Cupola for Excel.");
        try
        {
            loaded = Convert.ToBoolean(application.GetType().InvokeMember("RegisterXLL", BindingFlags.InvokeMethod, null, application, new object[] { xll }));
            if (!loaded) throw new InvalidOperationException();
        }
        catch { throw new COMException("Excel could not load Cupola. Check the installation and your organization's add-in policy."); }
    }
}
