using System;
using System.IO;
using Microsoft.Win32;

namespace QueryFarm.Vgi.ExcelDna;
internal static class UpdateEnvironment
{
    public static string CachePath(UpdateRelease release) => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QueryFarm", "Cupola", "Updates", release.Tag, UpdateRelease.InstallerName);
    public static bool DisabledByPolicy()
    {
        using var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var policy = machine.OpenSubKey(@"SOFTWARE\Policies\QueryFarm\Cupola");
        return Convert.ToInt32(policy?.GetValue("DisableUpdates", 0) ?? 0) != 0;
    }
    public static string InstalledVersion()
    {
        using var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var key = machine.OpenSubKey(@"SOFTWARE\QueryFarm\Cupola");
        return key?.GetValue("Version") as string ?? "";
    }
    public static string InstallDirectory()
    {
        using var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var key = machine.OpenSubKey(@"SOFTWARE\QueryFarm\Cupola");
        return key?.GetValue("InstallDirectory") as string ?? "";
    }
    public static bool SameDirectory(string a, string b) => !string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b) &&
        string.Equals(Path.GetFullPath(a).TrimEnd(Path.DirectorySeparatorChar), Path.GetFullPath(b).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
}
