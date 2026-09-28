using ExcelDna.Integration;

namespace QueryFarm.Vgi.ExcelDna;

public sealed class AddInLifecycle : IExcelAddIn
{
    public void AutoOpen()
    {
        SentryTelemetry.Initialize();
        try { ConnectionStore.ImportMachineDefaults(); }
        catch (System.Exception error) { ErrorLog.Write(error, "connections.import-defaults"); }
    }

    public void AutoClose()
    {
        try { HaybarnSessions.Cache.Dispose(); }
        finally { SentryTelemetry.Shutdown(); }
    }
}
