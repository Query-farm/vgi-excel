using System;
using System.IO;
using System.Text;
using System.Threading;

namespace QueryFarm.Vgi.ExcelDna;

internal static class ConnectionFile
{
    // A file lock coordinates independent Excel processes and logon sessions.
    internal static T Update<T>(string root, Func<T> update)
    {
        Directory.CreateDirectory(root);
        FileStream? gate = null;
        for (var attempt = 0; gate == null; attempt++)
        {
            try { gate = new FileStream(Path.Combine(root, "connections.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) when (attempt < 100) { Thread.Sleep(25); }
        }
        using (gate) return update();
    }

    internal static void Write(string path, string text)
    {
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                var bytes = new UTF8Encoding(false).GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(true);
            }
            for (var attempt = 0; ; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Replace(temporary, path, path + ".bak", true);
                    else File.Move(temporary, path);
                    break;
                }
                catch (IOException) when (attempt < 100) { Thread.Sleep(25); }
            }
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
