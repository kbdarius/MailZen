using System.IO;
using System.Text;
using System.Text.Json;

namespace EmailManage.Models;

/// <summary>
/// Captures what was learned from analysing the user's Deleted Items.
/// Stores sender/domain frequencies to give Ollama contextual hints during triage.
/// </summary>
public class LearnedProfile
{
    public string AccountKey { get; set; } = "";
    public DateTime LearnedAt { get; set; } = DateTime.UtcNow;
    public int TotalDeletedScanned { get; set; }

    /// <summary>Sender email → number of times found in Deleted Items.</summary>
    public Dictionary<string, int> DeletedSenderCounts { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Domain → number of times found in Deleted Items.</summary>
    public Dictionary<string, int> DeletedDomainCounts { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Senders the user explicitly kept (disagreed with AI). Skip these during triage.</summary>
    public HashSet<string> DoNotDeleteSenders { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Senders for which Outlook rules have already been created. Avoids duplicates.</summary>
    public HashSet<string> RuleCreatedSenders { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public void RegisterConfirmedKeep(string? senderEmail, string? domain)
    {
        var sender = (senderEmail ?? "").Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(sender))
        {
            DoNotDeleteSenders.Add(sender);
            DecrementCount(DeletedSenderCounts, sender);
        }

        var safeDomain = (domain ?? "").Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(safeDomain))
            DecrementCount(DeletedDomainCounts, safeDomain);
    }

    public void RegisterConfirmedDelete(string? senderEmail, string? domain)
    {
        var sender = (senderEmail ?? "").Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(sender))
        {
            IncrementCount(DeletedSenderCounts, sender);
            DoNotDeleteSenders.Remove(sender);
        }

        var safeDomain = (domain ?? "").Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(safeDomain))
            IncrementCount(DeletedDomainCounts, safeDomain);
    }

    // ── Persistence ──

    private static string GetProfilePath(string accountKey)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "learned_profile.json");
    }

    public void Save()
    {
        var path = GetProfilePath(AccountKey);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
    }

    public static LearnedProfile? Load(string accountKey)
    {
        var path = GetProfilePath(accountKey);
        if (!File.Exists(path)) return null;
        try
        {
            var json = File.ReadAllText(path);
            var profile = JsonSerializer.Deserialize<LearnedProfile>(json);
            return profile;
        }
        catch { return null; }
    }

    /// <summary>
    /// Generates a compact text summary of the user's deletion patterns.
    /// Injected into the Ollama prompt so the AI can consider personal history.
    /// </summary>
    public string ToPromptContext()
    {
        if (DeletedDomainCounts.Count == 0 && DeletedSenderCounts.Count == 0)
            return "";

        var sb = new StringBuilder();
        sb.AppendLine("USER DELETION HISTORY (from their Deleted Items folder):");

        // Top deleted domains (max 30)
        var topDomains = DeletedDomainCounts
            .OrderByDescending(kv => kv.Value)
            .Take(30)
            .ToList();

        if (topDomains.Count > 0)
        {
            sb.AppendLine("Frequently deleted domains:");
            foreach (var (domain, count) in topDomains)
                sb.AppendLine($"  - {domain} ({count}x)");
        }

        // Top deleted senders (max 20)
        var topSenders = DeletedSenderCounts
            .OrderByDescending(kv => kv.Value)
            .Take(20)
            .ToList();

        if (topSenders.Count > 0)
        {
            sb.AppendLine("Frequently deleted senders:");
            foreach (var (sender, count) in topSenders)
                sb.AppendLine($"  - {sender} ({count}x)");
        }

        sb.AppendLine($"Total deleted emails scanned: {TotalDeletedScanned}");
        sb.AppendLine("Emails from these senders/domains are more likely JUNK for this user.");

        // DO-NOT-DELETE list: senders the user explicitly protected
        if (DoNotDeleteSenders.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("DO NOT classify emails from these senders as JUNK (user explicitly kept them):");
            foreach (var sender in DoNotDeleteSenders.Take(50))
                sb.AppendLine($"  - {sender}");
        }

        return sb.ToString();
    }

    private static void IncrementCount(Dictionary<string, int> counts, string key)
    {
        if (counts.TryGetValue(key, out var current))
            counts[key] = current + 1;
        else
            counts[key] = 1;
    }

    private static void DecrementCount(Dictionary<string, int> counts, string key)
    {
        if (!counts.TryGetValue(key, out var current)) return;

        if (current <= 1)
            counts.Remove(key);
        else
            counts[key] = current - 1;
    }
}
