using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace EmailManage.Services;

/// <summary>Stores the user key in Windows Credential Manager; it never writes to SQLite or logs.</summary>
public sealed class OpenAiCredentialStore : ICredentialStore
{
    private const string Target = "MailZen/OpenAI";
    private const uint Generic = 1;
    private const uint PersistLocalMachine = 2;

    public Task SetApiKeyAsync(string apiKey, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(apiKey) || !apiKey.Trim().StartsWith("sk-", StringComparison.Ordinal)) throw new ArgumentException("Enter a valid OpenAI API key.", nameof(apiKey));
        var bytes = Encoding.UTF8.GetBytes(apiKey.Trim());
        var credential = new NativeCredential { Type = Generic, TargetName = Target, CredentialBlobSize = (uint)bytes.Length, Persist = PersistLocalMachine, UserName = "MailZen" };
        credential.CredentialBlob = Marshal.AllocHGlobal(bytes.Length);
        try
        {
            Marshal.Copy(bytes, 0, credential.CredentialBlob, bytes.Length);
            if (!CredWrite(ref credential, 0)) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        finally { Marshal.FreeHGlobal(credential.CredentialBlob); }
        return Task.CompletedTask;
    }

    public Task<bool> HasApiKeyAsync(CancellationToken cancellationToken = default) => Task.FromResult(TryReadApiKey() is not null);

    public Task RemoveApiKeyAsync(CancellationToken cancellationToken = default)
    { CredDelete(Target, Generic, 0); return Task.CompletedTask; }

    public string? TryReadApiKey()
    {
        if (!CredRead(Target, Generic, 0, out var pointer)) return null;
        try
        {
            var credential = Marshal.PtrToStructure<NativeCredential>(pointer);
            if (credential.CredentialBlob == IntPtr.Zero || credential.CredentialBlobSize == 0) return null;
            var bytes = new byte[credential.CredentialBlobSize];
            Marshal.Copy(credential.CredentialBlob, bytes, 0, bytes.Length);
            return Encoding.UTF8.GetString(bytes);
        }
        finally { CredFree(pointer); }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags, Type; public string TargetName; public string Comment; public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize; public IntPtr CredentialBlob; public uint Persist; public uint AttributeCount; public IntPtr Attributes; public string TargetAlias; public string UserName;
    }
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredWrite(ref NativeCredential userCredential, uint flags);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredDelete(string target, uint type, uint flags);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern void CredFree(IntPtr credential);
}

public sealed record AiPrivacySettings(bool LocalSearchOnly = false, int CandidateLimit = 20, decimal MonthlyBudgetUsd = 10m);
