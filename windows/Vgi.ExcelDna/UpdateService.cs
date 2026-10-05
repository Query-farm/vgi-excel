using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Win32;

namespace QueryFarm.Vgi.ExcelDna;
internal sealed class UpdateStatus
{
    public bool Supported { get; set; }
    public bool Daily { get; set; }
    public string Phase { get; set; } = "idle";
    public string Message { get; set; } = "Check for a newer Cupola release.";
    public string? LastChecked { get; set; }
    public string? Version { get; set; }
    public string? Build { get; set; }
    public string? Notes { get; set; }
    public UpdateStatus Copy() => (UpdateStatus)MemberwiseClone();
}
internal static class UpdateService
{
    private static readonly SemaphoreSlim Work = new(1, 1);
    private static readonly object Gate = new();
    private static readonly UpdateStatus State = new();
    private static UpdateRelease? Candidate;
    private const string Settings = @"Software\QueryFarm\Cupola\Updates";
    private static bool Supported => !UpdateEnvironment.DisabledByPolicy() && UpdateEnvironment.SameDirectory(UpdateEnvironment.InstallDirectory(), Path.GetDirectoryName(ExcelDnaUtil.XllPath)!);

    public static UpdateStatus Status()
    {
        lock (Gate)
        {
            State.Supported = Supported;
            using var settings = Registry.CurrentUser.OpenSubKey(Settings);
            State.Daily = Convert.ToInt32(settings?.GetValue("Daily", 1) ?? 1) != 0;
            State.LastChecked = settings?.GetValue("LastChecked") as string;
            if (!State.Supported) State.Message = UpdateEnvironment.DisabledByPolicy() ? "Updates are managed by your organization." : "Developer installation. Use Update Cupola for Excel.cmd from the new developer package.";
            return State.Copy();
        }
    }
    public static UpdateStatus Preference(bool daily)
    {
        using var settings = Registry.CurrentUser.CreateSubKey(Settings);
        settings.SetValue("Daily", daily ? 1 : 0, RegistryValueKind.DWord);
        return Status();
    }
    public static async Task<UpdateStatus> AutoCheck()
    {
        var state = Status();
        if (state.Supported && state.Daily && (!DateTime.TryParse(state.LastChecked, null, System.Globalization.DateTimeStyles.RoundtripKind, out var date) || DateTime.UtcNow - date.ToUniversalTime() > TimeSpan.FromDays(1)))
            return await Check();
        return state;
    }
    public static async Task<UpdateStatus> Check()
    {
        if (!Supported || !await Work.WaitAsync(0)) return Status();
        try
        {
            Set("checking", "Checking for updates…");
            using var settings = Registry.CurrentUser.CreateSubKey(Settings);
            settings.SetValue("LastChecked", DateTime.UtcNow.ToString("O"));
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            using var client = new UpdateClient();
            var release = await client.Latest(timeout.Token);
            lock (Gate)
            {
                Candidate = null;
                State.Version = State.Build = State.Notes = null;
                if (release.IsNewerThan(ProductInfo.Version, ProductInfo.Build))
                {
                    State.Version = release.Version; State.Build = release.Build; State.Notes = release.Notes;
                    if (release.CanUpgrade(ProductInfo.Version)) { Candidate = release; Set("available", "A new Cupola release is available."); }
                    else Set("manual", "This release uses the same installer version. See its release notes for installation instructions.");
                }
                else Set("current", "You’re up to date.");
            }
        }
        catch { Set("error", "Could not check for updates. Check your internet connection and try again."); }
        finally { Work.Release(); }
        return Status();
    }
    public static async Task<UpdateStatus> Download(Action<string>? progress)
    {
        if (!Supported || !await Work.WaitAsync(0)) return Status();
        try
        {
            var release = Candidate ?? throw new InvalidOperationException();
            Set("downloading", "Downloading and verifying the installer…");
            using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(10));
            using var client = new UpdateClient();
            await Task.Run(() => client.Download(release, UpdateEnvironment.CachePath(release), percent => progress?.Invoke("Downloading update… " + percent + "%"), timeout.Token));
            Set("ready", "The installer is downloaded and verified.");
        }
        catch { Set("error", "Could not download or verify the installer. No update was installed. Check for updates and try again."); }
        finally { Work.Release(); }
        return Status();
    }
    public static async Task<UpdateStatus> Install()
    {
        if (!Supported || !await Work.WaitAsync(0)) return Status();
        try
        {
            if (State.Phase != "ready" || Candidate is null) throw new InvalidOperationException("Download and verify the update first.");
            var helper = Path.Combine(UpdateEnvironment.InstallDirectory(), "Cupola.Updater.exe");
            // Run a copy so Windows Installer can replace the installed updater executable.
            var copy = Path.Combine(Path.GetDirectoryName(UpdateEnvironment.CachePath(Candidate))!, "Cupola.Updater-" + Guid.NewGuid().ToString("N") + ".exe");
            var release = Candidate;
            await Task.Run(() =>
            {
                File.Copy(helper, copy, false);
                using var locked = new FileStream(copy, FileMode.Open, FileAccess.Read, FileShare.Read);
                InstallerTrust.VerifyPublisher(copy);
                Process.Start(new ProcessStartInfo(copy, release.Version + " " + release.Build + " " + release.Sha256) { UseShellExecute = true, WorkingDirectory = UpdateEnvironment.InstallDirectory() });
            });
            Set("handoff", "The updater is open. Save your work and close every Excel window to continue. Reopen Excel after installation.");
        }
        catch { Set("error", "Could not open the updater. No update was installed. Check for updates and try again."); }
        finally { Work.Release(); }
        return Status();
    }
    private static void Set(string phase, string message) { lock (Gate) { State.Phase = phase; State.Message = message; } }
}
