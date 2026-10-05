using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using QueryFarm.Vgi.ExcelDna;

internal static class Program
{
    [STAThread] private static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        if (args.Length != 3 || !Regex.IsMatch(args[2], @"\A[0-9a-f]{64}\z")) return;
        var release = new UpdateRelease { Tag = "v" + args[0] + "-" + args[1], Sha256 = args[2] };
        try
        {
            if (!release.CanUpgrade(ProductInfo.Version) || !release.CanUpgrade(UpdateEnvironment.InstalledVersion()) || UpdateEnvironment.DisabledByPolicy())
                throw new InvalidOperationException();
            Application.Run(new UpdateWindow(release));
        }
        catch { MessageBox.Show("Open Settings → About in the installed Cupola add-in to start an update.", "Cupola for Excel", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }
}
internal sealed class UpdateWindow : Form
{
    private readonly UpdateRelease release;
    private readonly Label status = new() { Dock = DockStyle.Fill, AutoSize = false, Padding = new Padding(16), Text = "Save your work, then close every Excel window. This updater will wait; it will not close Excel or save your workbooks." };
    private readonly Button cancel = new() { Text = "Cancel", Dock = DockStyle.Bottom, Height = 40 };
    private readonly Timer timer = new() { Interval = 1500 };
    private bool installing;
    public UpdateWindow(UpdateRelease release)
    {
        this.release = release;
        Text = "Update Cupola for Excel"; ClientSize = new Size(420, 210); MinimumSize = new Size(360, 230); StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10); Controls.Add(status); Controls.Add(cancel);
        cancel.Click += (_, _) => Close();
        FormClosing += (_, e) => { if (installing) e.Cancel = true; };
        FormClosed += (_, _) => timer.Dispose();
        timer.Tick += async (_, _) => await TryInstall();
        Shown += (_, _) => timer.Start();
    }
    private async Task TryInstall()
    {
        if (installing) return;
        try
        {
            var processes = Process.GetProcessesByName("EXCEL");
            var open = processes.Length > 0;
            foreach (var process in processes) process.Dispose();
            if (open) { status.Text = "Waiting for Excel to close.\n\nSave your work, then close all Excel windows. You can cancel this update at any time before installation starts."; return; }
            timer.Stop(); installing = true; cancel.Enabled = false;
            status.Text = "Verifying Cupola " + release.Version + " (" + release.Build + ")…\n\nWindows may ask for administrator approval.";
            var code = await Task.Run(() =>
            {
                if (UpdateEnvironment.DisabledByPolicy()) throw new InvalidOperationException();
                var path = UpdateEnvironment.CachePath(release);
                // Keep the file locked against modification or replacement through installation.
                using var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                InstallerTrust.Verify(path, release);
                using var installer = Process.Start(new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "msiexec.exe"), "/i \"" + path + "\" /passive /norestart") { UseShellExecute = true, Verb = "runas" });
                if (installer is null) throw new InvalidOperationException();
                installer.WaitForExit(); return installer.ExitCode;
            });
            status.Text = code == 0 ? "Cupola was updated. Reopen Excel to use the new version." : code == 3010 ? "Cupola was updated. Restart Windows before opening Excel." : code == 1602 ? "Installation was cancelled. Your current Cupola installation is unchanged." : "Windows could not finish the update (code " + code + "). Close all Excel windows and try again from Cupola.";
        }
        catch { timer.Stop(); status.Text = "The update could not be verified or installed. No unverified installer was run. Reopen Cupola and check for updates to try again."; }
        finally { if (!timer.Enabled) { installing = false; cancel.Enabled = true; cancel.Text = "Close"; } }
    }
}
