using System.IO;
using System.Text.Json;
using EmailManage.Models;

namespace EmailManage.Services;

public class PatternAnalysisService
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;

    public PatternAnalysisService(OutlookConnectorService connector)
    {
        _connector = connector;
        _log = DiagnosticLogger.Instance;
    }

    public async Task<AnalysisResult> AnalyzeAccountAsync(OutlookAccountInfo account, IProgress<string>? progress = null)
    {
        var result = new AnalysisResult();
        try
        {
            progress?.Report("Fetching deleted items (last 12 months)...");
            var deletedEmails = await _connector.GetEmailsAsync(account.EmailAddress, account.StoreName, OutlookFolderType.DeletedItems, 12, 5000);
            result.DeletedCount = deletedEmails.Count;

            progress?.Report("Fetching kept inbox items (last 12 months)...");
            var keptEmails = await _connector.GetEmailsAsync(account.EmailAddress, account.StoreName, OutlookFolderType.Inbox, 12, 5000);
            
            // Apply keep grace period (e.g., older than 7 days)
            var cutoff = DateTime.Now.AddDays(-7);
            keptEmails = keptEmails.Where(e => e.ReceivedTime < cutoff).ToList();
            result.KeptCount = keptEmails.Count;

            progress?.Report("Extracting features and building rules...");
            var patterns = ExtractPatterns(deletedEmails, keptEmails);
            result.Patterns = patterns;

            progress?.Report("Saving pattern file...");
            SavePatternFile(account.AccountKey, patterns);

            progress?.Report("Saving state watermark...");
            SaveStateFile(account.AccountKey);

            result.Success = true;
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error analyzing account {Account}", account.EmailAddress);
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Re-scans recently deleted items (last N days) and merges new patterns into the existing pattern file.
    /// Uses a low threshold (1 delete) since the user is actively teaching the app.
    /// </summary>
    public async Task<LearnResult> LearnFromRecentDeletionsAsync(OutlookAccountInfo account, int days = 7, IProgress<string>? progress = null)
    {
        var result = new LearnResult();
        try
        {
            // 1. Load existing patterns
            progress?.Report("Loading existing patterns...");
            var existingPatterns = LoadExistingPatterns(account.AccountKey);
            var existingSenders = new HashSet<string>(existingPatterns.Where(p => p.Type == "Sender" || p.Type == "SenderWithUnsubscribe").Select(p => p.Value.ToLowerInvariant()));
            var existingDomains = new HashSet<string>(existingPatterns.Where(p => p.Type == "Domain").Select(p => p.Value.ToLowerInvariant()));

            // 2. Fetch recent deleted items
            progress?.Report($"Scanning deleted items from the last {days} days...");
            int months = Math.Max(1, (int)Math.Ceiling(days / 30.0));
            var deletedEmails = await _connector.GetEmailsAsync(account.EmailAddress, account.StoreName, OutlookFolderType.DeletedItems, months, 2000);
            
            // Filter to only recent items
            var cutoff = DateTime.Now.AddDays(-days);
            deletedEmails = deletedEmails.Where(e => e.ReceivedTime >= cutoff).ToList();
            result.RecentDeletedCount = deletedEmails.Count;
            progress?.Report($"Found {deletedEmails.Count} recently deleted emails...");

            // 3. Fetch current inbox to avoid false positives
            progress?.Report("Checking current inbox for safety...");
            var inboxEmails = await _connector.GetEmailsAsync(account.EmailAddress, account.StoreName, OutlookFolderType.Inbox, 3, 2000);
            var inboxSenders = new HashSet<string>(inboxEmails
                .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress))
                .Select(e => e.SenderEmailAddress.ToLowerInvariant()));

            // 4. Extract new sender patterns (threshold: just 1 delete since user is teaching)
            var newPatterns = new List<PatternRule>();
            var newSendersList = new List<string>();

            var recentSenders = deletedEmails
                .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress))
                .GroupBy(e => e.SenderEmailAddress.ToLowerInvariant())
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var kvp in recentSenders)
            {
                var sender = kvp.Key;
                var delCount = kvp.Value;

                // Skip if already in patterns
                if (existingSenders.Contains(sender)) continue;

                // Add as new pattern — user explicitly deleted these
                newPatterns.Add(new PatternRule
                {
                    RuleId = $"rule_sender_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                    Type = "Sender",
                    Value = sender,
                    Weight = 0.9,
                    EvidenceCount = delCount
                });
                newSendersList.Add(sender);
            }

            // 5. Extract new domain patterns (threshold: 2+ different senders from same domain)
            var getDomain = (string email) => {
                var parts = email.Split('@');
                return parts.Length == 2 ? parts[1] : "";
            };

            var commonProviders = new[] { "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "live.com", "icloud.com", "aol.com", "msn.com" };

            var recentDomains = deletedEmails
                .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress) && e.SenderEmailAddress.Contains("@"))
                .GroupBy(e => getDomain(e.SenderEmailAddress.ToLowerInvariant()))
                .Where(g => !commonProviders.Contains(g.Key) && !existingDomains.Contains(g.Key))
                .Where(g => g.Count() >= 2) // At least 2 emails from this domain
                .ToList();

            foreach (var group in recentDomains)
            {
                newPatterns.Add(new PatternRule
                {
                    RuleId = $"rule_domain_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                    Type = "Domain",
                    Value = group.Key,
                    Weight = 0.8,
                    EvidenceCount = group.Count()
                });
            }

            // 6. Merge and save
            if (newPatterns.Count > 0)
            {
                progress?.Report($"Adding {newPatterns.Count} new rules to pattern file...");
                var merged = existingPatterns.Concat(newPatterns).ToList();
                SavePatternFile(account.AccountKey, merged);
                SaveStateFile(account.AccountKey);
            }

            result.NewPatternsAdded = newPatterns.Count;
            result.NewSenders = newSendersList;
            result.TotalPatterns = existingPatterns.Count + newPatterns.Count;
            result.Success = true;

            progress?.Report($"Done! Added {newPatterns.Count} new rules. Total: {result.TotalPatterns} rules.");
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error learning from recent deletions for {Account}", account.EmailAddress);
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    private List<PatternRule> LoadExistingPatterns(string accountKey)
    {
        var patterns = new List<PatternRule>();
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailManage", "accounts", accountKey, "pattern_v1.yaml");
        if (!File.Exists(path)) return patterns;

        try
        {
            var lines = File.ReadAllLines(path);
            PatternRule? currentRule = null;

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("- id:"))
                {
                    if (currentRule != null) patterns.Add(currentRule);
                    currentRule = new PatternRule { RuleId = trimmed.Substring(5).Trim() };
                }
                else if (currentRule != null)
                {
                    if (trimmed.StartsWith("type:")) currentRule.Type = trimmed.Substring(5).Trim();
                    else if (trimmed.StartsWith("value:")) currentRule.Value = trimmed.Substring(6).Trim().Trim('"');
                    else if (trimmed.StartsWith("weight:"))
                    {
                        if (double.TryParse(trimmed.Substring(7).Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double w)) currentRule.Weight = w;
                    }
                    else if (trimmed.StartsWith("evidence_count:"))
                    {
                        if (int.TryParse(trimmed.Substring(15).Trim(), out int e)) currentRule.EvidenceCount = e;
                    }
                }
            }
            if (currentRule != null) patterns.Add(currentRule);
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to load existing pattern file");
        }

        return patterns;
    }

    private List<PatternRule> ExtractPatterns(List<EmailMessageInfo> deleted, List<EmailMessageInfo> kept)
    {
        var patterns = new List<PatternRule>();

        // 1. SENDER ANALYSIS
        var deletedSenders = deleted
            .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress))
            .GroupBy(e => e.SenderEmailAddress.ToLowerInvariant())
            .ToDictionary(g => g.Key, g => g.Count());

        var keptSenders = kept
            .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress))
            .GroupBy(e => e.SenderEmailAddress.ToLowerInvariant())
            .ToDictionary(g => g.Key, g => g.Count());

        foreach (var kvp in deletedSenders)
        {
            var sender = kvp.Key;
            var delCount = kvp.Value;
            var keepCount = keptSenders.TryGetValue(sender, out var k) ? k : 0;

            // Safety gate: Must have at least 5 deletes and 0 keeps to be a high-confidence delete rule
            if (delCount >= 5 && keepCount == 0)
            {
                patterns.Add(new PatternRule
                {
                    RuleId = $"rule_sender_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                    Type = "Sender",
                    Value = sender,
                    Weight = 0.9,
                    EvidenceCount = delCount
                });
            }
        }

        // 2. DOMAIN ANALYSIS
        var getDomain = (string email) => {
            var parts = email.Split('@');
            return parts.Length == 2 ? parts[1] : "";
        };

        var deletedDomains = deleted
            .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress) && e.SenderEmailAddress.Contains("@"))
            .GroupBy(e => getDomain(e.SenderEmailAddress.ToLowerInvariant()))
            .ToDictionary(g => g.Key, g => g.Count());

        var keptDomains = kept
            .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress) && e.SenderEmailAddress.Contains("@"))
            .GroupBy(e => getDomain(e.SenderEmailAddress.ToLowerInvariant()))
            .ToDictionary(g => g.Key, g => g.Count());

        var commonProviders = new[] { "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "live.com", "icloud.com", "aol.com", "msn.com" };

        foreach (var kvp in deletedDomains)
        {
            var domain = kvp.Key;
            if (string.IsNullOrWhiteSpace(domain) || commonProviders.Contains(domain)) continue;

            var delCount = kvp.Value;
            var keepCount = keptDomains.TryGetValue(domain, out var k) ? k : 0;

            // Domain rules need higher confidence (e.g., 10 deletes, 0 keeps)
            if (delCount >= 10 && keepCount == 0)
            {
                patterns.Add(new PatternRule
                {
                    RuleId = $"rule_domain_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                    Type = "Domain",
                    Value = domain,
                    Weight = 0.8,
                    EvidenceCount = delCount
                });
            }
        }

        // 3. SUBJECT ANALYSIS (Exact match for automated/recurring emails)
        var deletedSubjects = deleted
            .Where(e => !string.IsNullOrWhiteSpace(e.Subject) && e.Subject.Length > 5)
            .GroupBy(e => e.Subject.Trim())
            .ToDictionary(g => g.Key, g => g.Count());

        var keptSubjects = kept
            .Where(e => !string.IsNullOrWhiteSpace(e.Subject) && e.Subject.Length > 5)
            .GroupBy(e => e.Subject.Trim())
            .ToDictionary(g => g.Key, g => g.Count());

        foreach (var kvp in deletedSubjects)
        {
            var subject = kvp.Key;
            var delCount = kvp.Value;
            var keepCount = keptSubjects.TryGetValue(subject, out var k) ? k : 0;

            if (delCount >= 5 && keepCount == 0)
            {
                patterns.Add(new PatternRule
                {
                    RuleId = $"rule_subject_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                    Type = "Subject",
                    Value = subject,
                    Weight = 0.7,
                    EvidenceCount = delCount
                });
            }
        }

        // 4. BODY MARKERS (Newsletters/Promos with "unsubscribe")
        var deletedUnsubSenders = deleted
            .Where(e => !string.IsNullOrWhiteSpace(e.SenderEmailAddress) && 
                        !string.IsNullOrWhiteSpace(e.Body) && 
                        e.Body.Contains("unsubscribe", StringComparison.OrdinalIgnoreCase))
            .GroupBy(e => e.SenderEmailAddress.ToLowerInvariant())
            .ToDictionary(g => g.Key, g => g.Count());

        foreach (var kvp in deletedUnsubSenders)
        {
            var sender = kvp.Key;
            var delCount = kvp.Value;
            var keepCount = keptSenders.TryGetValue(sender, out var k) ? k : 0;

            // If it's a newsletter and we delete it often, it's a strong pattern
            if (delCount >= 3 && keepCount == 0)
            {
                if (!patterns.Any(p => p.Type == "Sender" && p.Value == sender))
                {
                    patterns.Add(new PatternRule
                    {
                        RuleId = $"rule_body_unsub_{Guid.NewGuid().ToString("N").Substring(0, 8)}",
                        Type = "SenderWithUnsubscribe",
                        Value = sender,
                        Weight = 0.95,
                        EvidenceCount = delCount
                    });
                }
            }
        }

        // Sort by evidence count descending
        return patterns.OrderByDescending(p => p.EvidenceCount).Take(50).ToList();
    }

    private void SavePatternFile(string accountKey, List<PatternRule> patterns)
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);

        var path = Path.Combine(dir, "pattern_v1.yaml");
        
        // Simple YAML generation
        using var writer = new StreamWriter(path);
        writer.WriteLine("version: 1");
        writer.WriteLine($"last_updated: {DateTime.UtcNow:O}");
        writer.WriteLine("rules:");
        foreach (var p in patterns)
        {
            writer.WriteLine($"  - id: {p.RuleId}");
            writer.WriteLine($"    type: {p.Type}");
            writer.WriteLine($"    value: \"{p.Value}\"");
            writer.WriteLine($"    weight: {p.Weight:F2}");
            writer.WriteLine($"    evidence_count: {p.EvidenceCount}");
        }
    }

    private void SaveStateFile(string accountKey)
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);

        var path = Path.Combine(dir, "state.json");
        var state = new
        {
            last_analyzed_at = DateTime.UtcNow,
            model_version = "1.0"
        };

        File.WriteAllText(path, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true }));
    }
}

public class AnalysisResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int DeletedCount { get; set; }
    public int KeptCount { get; set; }
    public List<PatternRule> Patterns { get; set; } = new();
}

public class PatternRule
{
    public string RuleId { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public double Weight { get; set; }
    public int EvidenceCount { get; set; }
}

public class LearnResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int RecentDeletedCount { get; set; }
    public int NewPatternsAdded { get; set; }
    public int TotalPatterns { get; set; }
    public List<string> NewSenders { get; set; } = new();
}
