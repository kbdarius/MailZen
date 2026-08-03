using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using EmailManage.Models;

namespace EmailManage.Services;

public sealed class OutlookOpenService : IOutlookOpenService
{
    public async Task<bool> TryOpenAsync(IndexedMessage message, CancellationToken cancellationToken = default) => await Task.Run(() =>
    {
        dynamic? app = null; dynamic? session = null; dynamic? item = null;
        try { app = GetActiveComObject("Outlook.Application"); session = app.GetNamespace("MAPI"); session.Logon("", "", false, false); item = session.GetItemFromID(message.EntryId, message.StoreId); item.Display(false); return true; }
        catch { return false; } finally { Release(item); Release(session); Release(app); }
    }, cancellationToken);

    public async Task<string> ExportToMsgAsync(IndexedMessage message, CancellationToken cancellationToken = default)
    {
        var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailZen", "Open Messages"); Directory.CreateDirectory(directory);
        var subject = string.IsNullOrWhiteSpace(message.Subject) ? "message" : message.Subject;
        var safe = string.Join("_", subject.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim(); safe = safe[..Math.Min(80, safe.Length)];
        var path = Path.Combine(directory, $"{DateTime.UtcNow:yyyyMMddHHmmss}_{safe}.msg");
        await Task.Run(() =>
        {
            dynamic? app = null; dynamic? session = null; dynamic? item = null;
            try { app = GetActiveComObject("Outlook.Application"); session = app.GetNamespace("MAPI"); session.Logon("", "", false, false); item = session.GetItemFromID(message.EntryId, message.StoreId); item.SaveAs(path, 9); }
            finally { Release(item); Release(session); Release(app); }
        }, cancellationToken);
        if (!File.Exists(path)) throw new InvalidOperationException("Outlook could not export this message.");
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); return path;
    }
    private static object GetActiveComObject(string progId) { var clsid = Type.GetTypeFromProgID(progId, true)!.GUID; var hr = GetActiveObject(ref clsid, IntPtr.Zero, out var obj); if (hr < 0) Marshal.ThrowExceptionForHR(hr); return obj; }
    [DllImport("oleaut32.dll", PreserveSig = true)] private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
