using System.IO;
using System.Text.Json;

namespace EmailManage.Models;

/// <summary>
/// Tracks Inbox emails tagged as MailZen: Temp, with their UTC expiry time.
/// Stored per-account. During Sync Review + Relearn, expired entries are
/// re-tagged as MailZen: Delete in Outlook automatically.
/// </summary>
public class TempRegistry
{
    public string AccountKey { get; set; } = "";

    /// <summary>EntryId → UTC expiry time.</summary>
    public Dictionary<string, DateTime> Entries { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    private static string GetPath(string accountKey)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "temp_registry.json");
    }

    public void Save()
    {
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(GetPath(AccountKey), json);
    }

    public static TempRegistry Load(string accountKey)
    {
        var path = GetPath(accountKey);
        if (!File.Exists(path)) return new TempRegistry { AccountKey = accountKey };
        try
        {
            var t = JsonSerializer.Deserialize<TempRegistry>(File.ReadAllText(path));
            return t ?? new TempRegistry { AccountKey = accountKey };
        }
        catch { return new TempRegistry { AccountKey = accountKey }; }
    }
}
