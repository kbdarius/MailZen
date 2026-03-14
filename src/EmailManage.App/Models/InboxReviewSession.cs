using System.IO;
using System.Text.Json;

namespace EmailManage.Models;

/// <summary>
/// Stores the latest inbox categorization run for one account so Outlook-side review
/// actions can later be synced back into the learned profile.
/// </summary>
public class InboxReviewSession
{
    public string AccountKey { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string StoreName { get; set; } = "";
    public string EmailAddress { get; set; } = "";
    public string DatasetCsvPath { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime SourceStartDate { get; set; }
    public DateTime SourceEndDate { get; set; }
    public int TotalInboxItems { get; set; }
    public int KeepCount { get; set; }
    public int ReviewCount { get; set; }
    public int DeleteCount { get; set; }
    public int TempCount { get; set; }
    public DateTime? LastSyncedAt { get; set; }
    public List<InboxReviewSessionItem> Items { get; set; } = new();

    private static string GetSessionPath(string accountKey)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "inbox_review_session.json");
    }

    public void Save()
    {
        var path = GetSessionPath(AccountKey);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
    }

    public static InboxReviewSession? Load(string accountKey)
    {
        var path = GetSessionPath(accountKey);
        if (!File.Exists(path)) return null;

        try
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<InboxReviewSession>(json);
        }
        catch
        {
            return null;
        }
    }
}

public class InboxReviewSessionItem
{
    public string EntryId { get; set; } = "";
    public string SenderEmail { get; set; } = "";
    public string Subject { get; set; } = "";
    public DateTime ReceivedTime { get; set; }
    public string Recommendation { get; set; } = "";
    public int JunkScore { get; set; }
}
