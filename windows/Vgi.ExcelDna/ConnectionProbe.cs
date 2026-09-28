using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace QueryFarm.Vgi.ExcelDna;

/// <summary>Disposable draft sessions; never enter the workbook session cache or connection store.</summary>
internal static class ConnectionProbe
{
    internal const string TimeoutMessage = "Connection test timed out after 20 seconds. Check the HTTPS address and your network, then try again.";
    private static readonly SemaphoreSlim Gate = new(1, 1);

    internal static async Task<T> Run<T>(Func<CancellationToken, T> operation, int timeoutMs = 20_000)
    {
        // A timed-out native call owns its handles until it exits. Bound the number
        // of outstanding probes rather than accumulating abandoned native sessions.
        if (!Gate.Wait(0)) throw new InvalidOperationException("The previous connection test is still stopping. Try again shortly.");
        var cancellation = new CancellationTokenSource();
        var task = Task.Run(() =>
        {
            try { return operation(cancellation.Token); }
            finally { Gate.Release(); }
        });
        using var delayCancellation = new CancellationTokenSource();
        var delay = Task.Delay(timeoutMs, delayCancellation.Token);
        if (await Task.WhenAny(task, delay) == task)
        {
            delayCancellation.Cancel();
            try { return await task; }
            finally { cancellation.Dispose(); }
        }
        cancellation.Cancel();
        // Observe late failures, and dispose only after the native query releases
        // its cancellation registration. Never free live native handles on timeout.
        _ = task.ContinueWith(completed => { _ = completed.Exception; cancellation.Dispose(); }, TaskScheduler.Default);
        throw new TimeoutException(TimeoutMessage);
    }

    internal static async Task<string[]> Catalogs(string location, Action<string>? progress = null)
    {
        var draft = new VgiConnection { Name = "discovery", Catalog = "discovery", Location = location };
        ConnectionStore.Validate(draft);
        async Task<string[]> Probe() => await Run(token =>
        {
            var credential = OAuthClient.IsSignedIn(draft) ? OAuthClient.GetAttachCredential(draft) : null;
            using var session = new NativeHaybarnSession(HaybarnClient.ProbePrelude(), token);
            var result = session.Query(CatalogScript(location, credential), null, token);
            return result.Rows.Select(row => row[0] as string).Where(value => !string.IsNullOrEmpty(value)).Cast<string>().Distinct().ToArray();
        });
        try { return await Probe(); }
        catch (Exception error) when (OAuthClient.ShouldPromptForSignIn(error))
        {
            progress?.Invoke("Waiting for sign-in…");
            try { await OAuthClient.SignInAsync(draft); }
            catch { throw new InvalidOperationException("Sign-in wasn’t completed. Try Find catalogs again, or enter the catalog name manually."); }
            progress?.Invoke("Finding catalogs…");
            return await Probe();
        }
    }

    internal static string CatalogScript(string location, OAuthAttachCredential? credential)
    {
        // Cupola owns credential storage. Never persist another copy in the engine.
        var option = credential is null ? "" : $", {credential.Option} := {HaybarnClient.SqlString(credential.Value)}";
        return $"SET vgi_oauth_enabled={(credential?.Option == "oauth_refresh_token" ? "true" : "false")};\n" +
               $"SELECT catalog FROM vgi_catalogs({HaybarnClient.SqlString(location)}, oauth_cache := 'none'{option});";
    }

    internal static async Task<QueryResult> Test(VgiConnection draft)
    {
        var members = ConnectionStore.ResolveAttachments(draft);
        foreach (var member in members) if (OAuthClient.IsSignedIn(member)) member.Authentication = "oauth";
        async Task<QueryResult> Probe() => await Run(token =>
        {
            using var session = new NativeHaybarnSession(HaybarnClient.ProbePrelude(), token);
            return session.Query(HaybarnClient.BuildSessionScript(members, "SELECT current_catalog(), current_schema();", probe: true), 1, token);
        });
        try { return await Probe(); }
        catch (Exception error) when (OAuthClient.ShouldPromptForSignIn(error))
        {
            // Identify members independently so a public catalog in a profile is
            // never sent through OAuth merely because another catalog requires it.
            foreach (var member in members)
            {
                if (members.Count > 1)
                {
                    try {
                        await Run(token => {
                            using var session = new NativeHaybarnSession(HaybarnClient.ProbePrelude(), token);
                            return session.Query(HaybarnClient.BuildSessionScript(new[] { member }, "SELECT 1;", probe: true), 1, token);
                        });
                        continue;
                    }
                    catch (Exception memberError) when (OAuthClient.ShouldPromptForSignIn(memberError)) { }
                }
                // Human sign-in runs outside the network deadline; only the
                // encrypted OAuth store changes, never the connection registry.
                await OAuthClient.SignInAsync(member);
                member.Authentication = "oauth";
            }
            return await Probe();
        }
    }
}
