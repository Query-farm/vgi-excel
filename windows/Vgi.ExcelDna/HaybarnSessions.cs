using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace QueryFarm.Vgi.ExcelDna;

internal interface IHaybarnSession : IDisposable
{
    QueryResult Query(string sql, int? maxRows, CancellationToken cancellation = default);
}

/// <summary>One serialized native session per friendly connection identity.</summary>
internal sealed class HaybarnSessionCache : IDisposable
{
    private sealed class Entry
    {
        internal readonly object Gate = new();
        internal string Fingerprint = "";
        internal IHaybarnSession? Session;
    }
    private readonly object gate = new();
    private readonly Dictionary<string, Entry> entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<string, CancellationToken, IHaybarnSession> create;
    private bool stopped;
    internal HaybarnSessionCache(Func<string, CancellationToken, IHaybarnSession> create) => this.create = create;

    internal QueryResult Query(string name, Func<string> prepare, string sql, int? maxRows, CancellationToken cancellation = default)
    {
        Entry entry;
        lock (gate)
        {
            if (stopped) throw new ObjectDisposedException(nameof(HaybarnSessionCache));
            if (!entries.TryGetValue(name, out entry!)) entries[name] = entry = new Entry();
        }
        cancellation.ThrowIfCancellationRequested();
        while (!Monitor.TryEnter(entry.Gate, 50)) cancellation.ThrowIfCancellationRequested();
        try
        {
            cancellation.ThrowIfCancellationRequested();
            lock (gate) if (stopped) throw new ObjectDisposedException(nameof(HaybarnSessionCache));
            try
            {
                // Resolve settings and credentials after acquiring the session lock.
                // Retain only a digest, never a copy of the credential-bearing ATTACH.
                var setup = prepare();
                string fingerprint;
                using (var hash = SHA256.Create()) fingerprint = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(setup)));
                if (entry.Session is null || entry.Fingerprint != fingerprint)
                {
                    entry.Session?.Dispose();
                    entry.Session = null;
                    cancellation.ThrowIfCancellationRequested();
                    entry.Session = create(setup, cancellation);
                    entry.Fingerprint = fingerprint;
                }
                return entry.Session.Query(sql, maxRows, cancellation);
            }
            catch (OperationCanceledException) when (entry.Session is not null)
            {
                // An interrupted statement does not invalidate the native connection.
                throw;
            }
            catch
            {
                // Do not replay arbitrary SQL: it might already have had effects.
                entry.Session?.Dispose();
                entry.Session = null;
                entry.Fingerprint = "";
                throw;
            }
        }
        finally { Monitor.Exit(entry.Gate); }
    }

    internal void Invalidate(string name, Action? change = null)
    {
        Entry entry;
        lock (gate)
        {
            if (!entries.TryGetValue(name, out entry!)) entries[name] = entry = new Entry();
        }
        lock (entry.Gate)
        {
            entry.Session?.Dispose(); entry.Session = null; entry.Fingerprint = "";
            change?.Invoke();
        }
    }

    public void Dispose()
    {
        Entry[] snapshot;
        lock (gate)
        {
            stopped = true;
            snapshot = new List<Entry>(entries.Values).ToArray();
            entries.Clear();
        }
        foreach (var entry in snapshot)
            lock (entry.Gate) { entry.Session?.Dispose(); entry.Session = null; }
    }
}

internal static class HaybarnSessions
{
    internal static readonly HaybarnSessionCache Cache = new((setup, cancellation) => new NativeHaybarnSession(setup, cancellation));
}
