using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace QueryFarm.Vgi.ExcelDna;

internal static class InstallerTrust
{
    public static void Verify(string path, UpdateRelease release)
    {
        using (var stream = File.OpenRead(path))
        using (var sha = SHA256.Create())
            if (!string.Equals(BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", ""), release.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The installer checksum does not match the release.");
        VerifyPublisher(path);
        VerifyIdentity(path, release);
    }
    internal static void VerifyIdentity(string path, UpdateRelease release)
    {
        uint database = 0;
        try
        {
            if (MsiOpenDatabase(path, IntPtr.Zero, out database) != 0) throw new InvalidDataException("The installer could not be inspected.");
            if (Property(database, "UpgradeCode") != UpdateRelease.UpgradeCode || Property(database, "ProductName") != "Cupola for Excel" ||
                Property(database, "Manufacturer") != "Query Farm LLC" || Property(database, "ProductVersion") != release.Version || Property(database, "CUPOLABUILD") != release.Build)
                throw new InvalidDataException("The installer identity does not match this Cupola release.");
        }
        finally { if (database != 0) MsiCloseHandle(database); }
    }
    private static string Property(uint database, string name)
    {
        uint view = 0, record = 0;
        try
        {
            if (MsiDatabaseOpenView(database, "SELECT `Value` FROM `Property` WHERE `Property` = '" + name + "'", out view) != 0 || MsiViewExecute(view, 0) != 0 || MsiViewFetch(view, out record) != 0)
                return "";
            var value = new StringBuilder(256); uint length = 256;
            return MsiRecordGetString(record, 1, value, ref length) == 0 ? value.ToString() : "";
        }
        finally { if (record != 0) MsiCloseHandle(record); if (view != 0) MsiCloseHandle(view); }
    }
    internal static void VerifyPublisher(string path)
    {
        var file = new TrustFile { Size = (uint)Marshal.SizeOf(typeof(TrustFile)), Path = path };
        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(TrustFile)));
        Marshal.StructureToPtr(file, pointer, false);
        var data = new TrustData { Size = (uint)Marshal.SizeOf(typeof(TrustData)), UiChoice = 2, Revocation = 1, UnionChoice = 1, File = pointer, StateAction = 1, ProviderFlags = 0x80 };
        var action = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
        try
        {
            if (WinVerifyTrust(new IntPtr(-1), ref action, ref data) != 0) throw new InvalidDataException("Windows could not verify the installer signature. Check your internet connection and try again.");
            var provider = WTHelperProvDataFromStateData(data.State);
            var signerPointer = WTHelperGetProvSignerFromChain(provider, 0, false, 0);
            if (signerPointer == IntPtr.Zero) throw new InvalidDataException("The installer has no verified publisher.");
            var signer = Marshal.PtrToStructure<ProviderSigner>(signerPointer);
            if (signer.CertCount == 0 || signer.CertChain == IntPtr.Zero) throw new InvalidDataException("The installer publisher is missing.");
            var certificate = Marshal.PtrToStructure<ProviderCert>(signer.CertChain);
            using var publisher = new X509Certificate2(certificate.Certificate);
            if (!string.Equals(publisher.Subject, "CN=Query Farm LLC, O=Query Farm LLC, L=Glen Allen, S=Virginia, C=US", StringComparison.Ordinal))
                throw new InvalidDataException("The installer was not signed by Query Farm LLC.");
        }
        finally
        {
            data.StateAction = 2;
            WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            Marshal.DestroyStructure<TrustFile>(pointer); Marshal.FreeHGlobal(pointer);
        }
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct TrustFile { public uint Size; [MarshalAs(UnmanagedType.LPWStr)] public string Path; public IntPtr Handle, Subject; }
    [StructLayout(LayoutKind.Sequential)] private struct TrustData
    {
        public uint Size; public IntPtr Policy, Sip; public uint UiChoice, Revocation, UnionChoice; public IntPtr File; public uint StateAction; public IntPtr State, Url; public uint ProviderFlags, UiContext; public IntPtr SignatureSettings;
    }
    [StructLayout(LayoutKind.Sequential)] private struct ProviderSigner { public uint Size; public System.Runtime.InteropServices.ComTypes.FILETIME VerifiedAt; public uint CertCount; public IntPtr CertChain; }
    [StructLayout(LayoutKind.Sequential)] private struct ProviderCert { public uint Size; public IntPtr Certificate; }
    [DllImport("wintrust.dll", ExactSpelling = true)] private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);
    [DllImport("wintrust.dll", ExactSpelling = true)] private static extern IntPtr WTHelperProvDataFromStateData(IntPtr state);
    [DllImport("wintrust.dll", ExactSpelling = true)] private static extern IntPtr WTHelperGetProvSignerFromChain(IntPtr provider, uint index, [MarshalAs(UnmanagedType.Bool)] bool counterSigner, uint counterIndex);
    [DllImport("msi.dll", CharSet = CharSet.Unicode, EntryPoint = "MsiOpenDatabaseW")] private static extern uint MsiOpenDatabase(string path, IntPtr mode, out uint database);
    [DllImport("msi.dll", CharSet = CharSet.Unicode, EntryPoint = "MsiDatabaseOpenViewW")] private static extern uint MsiDatabaseOpenView(uint database, string query, out uint view);
    [DllImport("msi.dll")] private static extern uint MsiViewExecute(uint view, uint record);
    [DllImport("msi.dll")] private static extern uint MsiViewFetch(uint view, out uint record);
    [DllImport("msi.dll", CharSet = CharSet.Unicode, EntryPoint = "MsiRecordGetStringW")] private static extern uint MsiRecordGetString(uint record, uint field, StringBuilder value, ref uint length);
    [DllImport("msi.dll")] private static extern uint MsiCloseHandle(uint handle);
}
