using System;
using System.IO;
using System.Text;
using QueryFarm.Vgi.ExcelDna;

internal static class UpdaterTests
{
    internal static void Run()
    {
        const string tag = "v0.5.2-20261001.1";
        string Asset(string name, string? url = null) => "{\"name\":\"" + name + "\",\"size\":123,\"browser_download_url\":\"" + (url ?? "https://github.com/Query-farm/vgi-excel/releases/download/" + tag + "/" + name) + "\"}";
        string Json(string assets, string extra = "") => "{\"tag_name\":\"" + tag + "\",\"draft\":false,\"prerelease\":false,\"assets\":[" + assets + "]" + extra + "}";
        var valid = Json(Asset("CupolaForExcel.msi") + "," + Asset("SHA256SUMS.txt"));
        var release = UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid));
        Ensure(release.Version == "0.5.2" && release.Build == "20261001.1", "version/build parsed");
        Ensure(release.IsNewerThan("0.5.1", "20260930.4") && release.CanUpgrade("0.5.1"), "patch upgrade");
        Ensure(!release.IsNewerThan("0.6.0", "20260901.1") && !release.CanUpgrade("0.6.0"), "no downgrade");
        Ensure(release.IsNewerThan("0.5.2", "20261001.0") && !release.CanUpgrade("0.5.2"), "same MSI version needs manual installation");
        Ensure(!release.IsNewerThan("0.5.2", "20261001.1"), "current release");
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid.Replace("\"draft\":false", "\"draft\":true"))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid.Replace("\"prerelease\":false", "\"prerelease\":true"))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid.Replace(tag, "v0.5.2-preview"))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid.Replace("\"size\":123", "\"size\":0"))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(valid.Replace("\"size\":123", "\"size\":9999999999"))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(Json(Asset("CupolaForExcel.msi", "https://evil.example/installer.msi") + "," + Asset("SHA256SUMS.txt")))));
        Reject(() => UpdateRelease.Parse(Encoding.UTF8.GetBytes(Json(Asset("CupolaForExcel.msi") + "," + Asset("CupolaForExcel.msi") + "," + Asset("SHA256SUMS.txt")))));
        var hash = new string('a', 64);
        Ensure(UpdateRelease.ReadChecksum(hash + "  CupolaForExcel.msi\r\n") == hash, "checksum");
        Reject(() => UpdateRelease.ReadChecksum(hash + "  CupolaForExcel.msi\n" + hash + "  CupolaForExcel.msi\n"));
        Reject(() => UpdateRelease.ReadChecksum(hash + "  another.msi"));
        foreach (var url in new[] { "http://github.com/a", "https://github.com.evil.example/a", "https://user@github.com/a", "https://github.com:444/a", "https://github.com/a#fragment", "file:///tmp/a", "https://localhost/a" })
            Ensure(!UpdateRelease.AllowedDownload(new Uri(url)), "reject download address");
        Ensure(UpdateRelease.AllowedDownload(new Uri("https://release-assets.githubusercontent.com/a?signature=synthetic")), "GitHub asset redirect");
        var path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, "not an installer"); release.Sha256 = hash;
            Reject(() => InstallerTrust.Verify(path, release));
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                Reject(() => InstallerTrust.VerifyPublisher(path));
            else Console.WriteLine("SKIP: Windows Authenticode rejection requires Windows.");
        }
        finally { File.Delete(path); }
        Console.WriteLine("PASS: update release, origin, version, and checksum policy");
    }
    private static void Ensure(bool value, string description) { if (!value) throw new Exception("Updater assertion failed: " + description); }
    private static void Reject(Action action) { try { action(); } catch (InvalidDataException) { return; } throw new Exception("Unsafe update was accepted."); }
}
