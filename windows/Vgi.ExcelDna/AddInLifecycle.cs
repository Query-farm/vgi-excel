using ExcelDna.Integration;

namespace QueryFarm.Vgi.ExcelDna;

public sealed class AddInLifecycle : IExcelAddIn
{
    public void AutoOpen() => SentryTelemetry.Initialize();

    public void AutoClose() => SentryTelemetry.Shutdown();
}
