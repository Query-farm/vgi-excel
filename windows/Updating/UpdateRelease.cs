using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace QueryFarm.Vgi.ExcelDna;

[DataContract]
internal sealed class UpdateRelease
{
    [DataMember(Name = "tag_name")] public string Tag = "";
    [DataMember(Name = "draft")] public bool Draft;
    [DataMember(Name = "prerelease")] public bool Prerelease;
    [DataMember(Name = "assets")] public UpdateAsset[] Assets = Array.Empty<UpdateAsset>();
    public string Version => Parts().Groups[1].Value;
    public string Build => Parts().Groups[2].Value;
    public string Notes => "https://github.com/Query-farm/vgi-excel/releases/tag/" + Tag;
    public string Sha256 = "";
    public const string InstallerName = "CupolaForExcel.msi";
    public const string UpgradeCode = "{9C26AB13-1E7D-4D5F-948D-BBDDC9444346}";
    public const long MaximumBytes = 512L * 1024 * 1024;

    private Match Parts()
    {
        var match = Regex.Match(Tag ?? "", @"\Av(\d{1,3}\.\d{1,3}\.\d{1,5})-(\d{8}\.\d{1,5})\z");
        if (!match.Success || !System.Version.TryParse(match.Groups[1].Value, out var version) || version.Major > 255 || version.Minor > 255 || version.Build > 65535)
            throw new InvalidDataException("The update release has an invalid version.");
        return match;
    }
    public static UpdateRelease Parse(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        var release = (UpdateRelease)new DataContractJsonSerializer(typeof(UpdateRelease)).ReadObject(stream)!;
        _ = release.Parts();
        if (release.Draft || release.Prerelease) throw new InvalidDataException("Only published stable releases can be installed.");
        release.Asset(InstallerName);
        release.Asset("SHA256SUMS.txt");
        return release;
    }
    public UpdateAsset Asset(string name)
    {
        var matches = (Assets ?? Array.Empty<UpdateAsset>()).Where(x => x.Name == name).ToArray();
        var expected = "https://github.com/Query-farm/vgi-excel/releases/download/" + Tag + "/" + name;
        if (matches.Length != 1 || matches[0].Url != expected || matches[0].Size <= 0 || matches[0].Size > MaximumBytes)
            throw new InvalidDataException("The release does not contain a supported installer.");
        return matches[0];
    }
    public bool IsNewerThan(string version, string build)
    {
        var comparison = System.Version.Parse(Version).CompareTo(System.Version.Parse(version));
        return comparison > 0 || comparison == 0 && System.Version.Parse(Build).CompareTo(System.Version.Parse(build)) > 0;
    }
    public bool CanUpgrade(string version) => System.Version.Parse(Version) > System.Version.Parse(version);
    public static string ReadChecksum(string text)
    {
        var matches = Regex.Matches(text, @"(?m)^([0-9a-fA-F]{64})[ \t]+\*?CupolaForExcel\.msi\r?$");
        if (matches.Count != 1) throw new InvalidDataException("The installer checksum is missing or ambiguous.");
        return matches[0].Groups[1].Value.ToLowerInvariant();
    }
    public static bool AllowedDownload(Uri url) => url.Scheme == "https" && url.IsDefaultPort && url.UserInfo.Length == 0 && url.Fragment.Length == 0 &&
        new[] { "api.github.com", "github.com", "release-assets.githubusercontent.com", "objects.githubusercontent.com" }.Contains(url.DnsSafeHost, StringComparer.OrdinalIgnoreCase);
}
[DataContract]
internal sealed class UpdateAsset
{
    [DataMember(Name = "name")] public string Name = "";
    [DataMember(Name = "browser_download_url")] public string Url = "";
    [DataMember(Name = "size")] public long Size;
}
