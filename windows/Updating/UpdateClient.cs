using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class UpdateClient : IDisposable
{
    private readonly HttpClient client = new(new HttpClientHandler { AllowAutoRedirect = false });
    public UpdateClient()
    {
        client.Timeout = Timeout.InfiniteTimeSpan;
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Cupola-for-Excel/" + ProductInfo.Version);
    }
    public async Task<UpdateRelease> Latest(CancellationToken token)
    {
        var release = UpdateRelease.Parse(await Bytes("https://api.github.com/repos/Query-farm/vgi-excel/releases/latest", 1024 * 1024, token));
        release.Sha256 = UpdateRelease.ReadChecksum(Encoding.UTF8.GetString(await Bytes(release.Asset("SHA256SUMS.txt").Url, 1024 * 1024, token)));
        return release;
    }
    private async Task<HttpResponseMessage> Open(string address, CancellationToken token)
    {
        var uri = new Uri(address);
        for (var redirects = 0; redirects < 6; redirects++)
        {
            if (!UpdateRelease.AllowedDownload(uri)) throw new InvalidDataException("The update download address is not trusted.");
            var response = await client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
            if ((int)response.StatusCode >= 300 && (int)response.StatusCode <= 399)
            {
                var target = response.Headers.Location;
                response.Dispose();
                if (target is null) throw new InvalidDataException("The update download is unavailable.");
                uri = target.IsAbsoluteUri ? target : new Uri(uri, target);
                continue;
            }
            if (!response.IsSuccessStatusCode) { response.Dispose(); throw new IOException("The update service is unavailable. Try again later."); }
            return response;
        }
        throw new IOException("The update download redirected too many times.");
    }
    private async Task<byte[]> Bytes(string url, long limit, CancellationToken token)
    {
        using var result = new MemoryStream();
        await Copy(url, result, limit, null, token).ConfigureAwait(false);
        return result.ToArray();
    }
    private async Task<long> Copy(string url, Stream destination, long limit, Action<int>? progress, CancellationToken token)
    {
        using var response = await Open(url, token).ConfigureAwait(false);
        if (response.Content.Headers.ContentLength > limit) throw new InvalidDataException("The update download is too large.");
        using var source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
        var buffer = new byte[81920]; long total = 0; int last = -1;
        while (true)
        {
            var count = await source.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false);
            if (count == 0) break;
            total += count;
            if (total > limit) throw new InvalidDataException("The update download is too large.");
            await destination.WriteAsync(buffer, 0, count, token).ConfigureAwait(false);
            var percent = (int)(total * 100 / limit);
            if (percent != last) { last = percent; progress?.Invoke(percent); }
        }
        return total;
    }
    public async Task Download(UpdateRelease release, string path, Action<int> progress, CancellationToken token)
    {
        var asset = release.Asset(UpdateRelease.InstallerName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".partial.msi";
        try
        {
            using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                if (await Copy(asset.Url, output, asset.Size, progress, token).ConfigureAwait(false) != asset.Size) throw new InvalidDataException("The installer download is incomplete.");
            InstallerTrust.Verify(temporary, release);
            if (File.Exists(path)) File.Delete(path);
            File.Move(temporary, path);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
    public void Dispose() => client.Dispose();
}
