using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace QueryFarm.Vgi.ExcelDna;

internal static class CupolaWindowIcon
{
    internal static void Apply(Form window)
    {
        // Use the embedded mark so branding also works before WebView starts and
        // in the native fallback, without depending on loose web assets.
        using var stream = typeof(CupolaWindowIcon).Assembly.GetManifestResourceStream("QueryFarm.Vgi.ExcelDna.OpeningMark")
            ?? throw new InvalidOperationException("The embedded Cupola mark is missing.");
        using var image = Image.FromStream(stream);
        using var bitmap = new Bitmap(image);
        var handle = bitmap.GetHicon();
        try
        {
            using var borrowed = Icon.FromHandle(handle);
            var icon = (Icon)borrowed.Clone();
            window.Icon = icon;
            window.ShowIcon = true;
            window.Disposed += (_, __) => icon.Dispose();
        }
        finally { DestroyIcon(handle); }
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr icon);
}
