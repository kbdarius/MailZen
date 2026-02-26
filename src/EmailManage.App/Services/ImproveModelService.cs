using System.Globalization;
using System.IO;
using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Handles the "Improve Model" workflow: loading review items, processing user actions,
/// updating patterns, logging feedback, and supporting undo.
/// </summary>
public class ImproveModelService
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;

    /// <summary>
    /// Tracks the last action batch for undo.
    /// </summary>
    public UndoBatch? LastBatch { get; private set; }

    public ImproveModelService(OutlookConnectorService connector)
    {
        _connector = connector;
        _log = DiagnosticLogger.Instance;
    }

    /// <summary>
    /// Loads emails from the "Needs Review" folder, scores them with current patterns, and groups by domain.
    /// </summary>
    public async Task<(List<ReviewEmailItem> Items, List<ReviewGroup> Groups)> LoadNeedsReviewAsync(
        OutlookAccountInfo account, IProgress<string>? progress = null)
    {
        progress?.Report("Loading Needs Review emails...");
        var rawEmails = await _connector.GetTriageFolderEmailsAsync(
            account.EmailAddress, account.StoreName,
            OutlookConnectorService.NeedsReviewFolderName, 500);

        progress?.Report($"Found {rawEmails.Count} emails. Scoring...");

        // Load patterns to score each item
        var patterns = LoadPatterns(account.AccountKey);

        var items = new List<ReviewEmailItem>();
        foreach (var email in rawEmails)
        {
            var sender = email.SenderEmailAddress?.ToLowerInvariant() ?? "";
            var domain = sender.Contains("@") ? sender.Split('@')[1] : "";
            var score = ScoreEmail(email, patterns);

            var band = score.Confidence >= 0.85 ? "High" :
                       score.Confidence >= 0.7 ? "Medium" : "Low";

            items.Add(new ReviewEmailItem
            {
                EntryId = email.EntryId,
                Subject = email.Subject,
                SenderName = email.SenderName,
                SenderEmail = sender,
                Domain = domain,
                ReceivedTime = email.ReceivedTime,
                Confidence = score.Confidence,
                ConfidenceBand = band,
                MatchedRuleIds = score.MatchedRules.Select(r => r.RuleId).ToList(),
                MatchedRulesSummary = string.Join(", ", score.MatchedRules.Select(r => $"{r.Type}:{r.Value}"))
            });
        }

        // Group by domain
        var groups = items
            .GroupBy(i => string.IsNullOrWhiteSpace(i.Domain) ? "(unknown)" : i.Domain)
            .Select(g => new ReviewGroup
            {
                GroupKey = g.Key,
                GroupLabel = g.Key,
                ItemCount = g.Count(),
                AvgConfidence = Math.Round(g.Average(i => i.Confidence), 2),
                Items = g.OrderByDescending(i => i.Confidence).ToList()
            })
            .OrderByDescending(g => g.AvgConfidence)
            .ThenByDescending(g => g.ItemCount)
            .ToList();

        progress?.Report($"Loaded {items.Count} items in {groups.Count} groups.");
        return (items, groups);
    }

    /// <summary>
    /// Processes selected items: Delete + Learn — deletes emails AND updates patterns.
    /// </summary>
    public async Task<ImproveActionResult> DeleteAndLearnAsync(
        OutlookAccountInfo account, List<ReviewEmailItem> selectedItems, IProgress<string>? progress = null)
    {
        var result = new ImproveActionResult { Action = "Delete + Learn" };
        var undoEntries = new List<UndoEntry>();

        try
        {
            int processed = 0;
            foreach (var item in selectedItems)
            {
                progress?.Report($"Deleting {processed + 1}/{selectedItems.Count}: {item.SenderName}...");
                var newId = await _connector.MoveToDeletedItemsAsync(item.EntryId, account.EmailAddress, account.StoreName);
                if (newId != null)
                {
                    undoEntries.Add(new UndoEntry
                    {
                        EntryId = newId,
                        OriginalFolderName = item.OriginalFolderName,
                        DestinationFolderName = "DeletedItems",
                        SenderEmail = item.SenderEmail,
                        Subject = item.Subject
                    });
                    processed++;
                }
            }

            // Learn from these deletions
            progress?.Report("Updating patterns from your decisions...");
            var (patternsAdded, patternsUpdated) = UpdatePatternsFromFeedback(account.AccountKey, selectedItems, "delete");

            // Log feedback
            LogFeedback(account.AccountKey, selectedItems, "Delete+Learn");

            result.ProcessedCount = processed;
            result.NewPatternsAdded = patternsAdded;
            result.PatternsUpdated = patternsUpdated;
            result.AffectedSenders = selectedItems.Select(i => i.SenderEmail).Distinct().ToList();
            result.Success = true;

            LastBatch = new UndoBatch
            {
                Timestamp = DateTime.UtcNow,
                Action = "Delete+Learn",
                Entries = undoEntries
            };
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error in DeleteAndLearn");
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Processes selected items: Delete Only — deletes emails without learning.
    /// </summary>
    public async Task<ImproveActionResult> DeleteOnlyAsync(
        OutlookAccountInfo account, List<ReviewEmailItem> selectedItems, IProgress<string>? progress = null)
    {
        var result = new ImproveActionResult { Action = "Delete Only" };
        var undoEntries = new List<UndoEntry>();

        try
        {
            int processed = 0;
            foreach (var item in selectedItems)
            {
                progress?.Report($"Deleting {processed + 1}/{selectedItems.Count}: {item.SenderName}...");
                var newId = await _connector.MoveToDeletedItemsAsync(item.EntryId, account.EmailAddress, account.StoreName);
                if (newId != null)
                {
                    undoEntries.Add(new UndoEntry
                    {
                        EntryId = newId,
                        OriginalFolderName = item.OriginalFolderName,
                        DestinationFolderName = "DeletedItems",
                        SenderEmail = item.SenderEmail,
                        Subject = item.Subject
                    });
                    processed++;
                }
            }

            LogFeedback(account.AccountKey, selectedItems, "DeleteOnly");

            result.ProcessedCount = processed;
            result.Success = true;

            LastBatch = new UndoBatch
            {
                Timestamp = DateTime.UtcNow,
                Action = "DeleteOnly",
                Entries = undoEntries
            };
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error in DeleteOnly");
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Processes selected items: Keep + Learn — moves emails back to Inbox AND learns to keep them.
    /// </summary>
    public async Task<ImproveActionResult> KeepAndLearnAsync(
        OutlookAccountInfo account, List<ReviewEmailItem> selectedItems, IProgress<string>? progress = null)
    {
        var result = new ImproveActionResult { Action = "Keep + Learn" };
        var undoEntries = new List<UndoEntry>();

        try
        {
            int processed = 0;
            foreach (var item in selectedItems)
            {
                progress?.Report($"Keeping {processed + 1}/{selectedItems.Count}: {item.SenderName}...");
                var newId = await _connector.MoveToInboxAsync(item.EntryId, account.EmailAddress, account.StoreName);
                if (newId != null)
                {
                    undoEntries.Add(new UndoEntry
                    {
                        EntryId = newId,
                        OriginalFolderName = item.OriginalFolderName,
                        DestinationFolderName = "Inbox",
                        SenderEmail = item.SenderEmail,
                        Subject = item.Subject
                    });
                    processed++;
                }
            }

            // Learn to keep: reduce weight for matched rules or add keep-intent markers
            progress?.Report("Updating patterns to protect these senders...");
            var (patternsAdded, patternsUpdated) = UpdatePatternsFromFeedback(account.AccountKey, selectedItems, "keep");

            LogFeedback(account.AccountKey, selectedItems, "Keep+Learn");

            result.ProcessedCount = processed;
            result.NewPatternsAdded = patternsAdded;
            result.PatternsUpdated = patternsUpdated;
            result.AffectedSenders = selectedItems.Select(i => i.SenderEmail).Distinct().ToList();
            result.Success = true;

            LastBatch = new UndoBatch
            {
                Timestamp = DateTime.UtcNow,
                Action = "Keep+Learn",
                Entries = undoEntries
            };
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error in KeepAndLearn");
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Undoes the last batch by moving emails back to their original folders.
    /// </summary>
    public async Task<ImproveActionResult> UndoLastBatchAsync(
        OutlookAccountInfo account, IProgress<string>? progress = null)
    {
        var result = new ImproveActionResult { Action = "Undo" };

        if (LastBatch == null || LastBatch.Entries.Count == 0)
        {
            result.ErrorMessage = "No batch to undo.";
            return result;
        }

        try
        {
            int processed = 0;
            foreach (var entry in LastBatch.Entries)
            {
                progress?.Report($"Undoing {processed + 1}/{LastBatch.Entries.Count}: {entry.SenderEmail}...");

                string? newId = null;
                if (entry.DestinationFolderName == "DeletedItems")
                {
                    // Was deleted — move back to Needs Review
                    newId = await _connector.MoveToTriageFolderAsync(
                        entry.EntryId, account.EmailAddress, account.StoreName,
                        OutlookConnectorService.NeedsReviewFolderName);
                }
                else if (entry.DestinationFolderName == "Inbox")
                {
                    // Was kept — move back to Needs Review
                    newId = await _connector.MoveToTriageFolderAsync(
                        entry.EntryId, account.EmailAddress, account.StoreName,
                        OutlookConnectorService.NeedsReviewFolderName);
                }

                if (newId != null) processed++;
            }

            result.ProcessedCount = processed;
            result.Success = true;

            // If this was a +Learn action, we should also revert pattern changes
            // For now, we note this in the log
            if (LastBatch.Action.Contains("Learn"))
            {
                _log.Warn("Undo of a +Learn action does not revert pattern changes. Consider re-running analysis.");
            }

            LogFeedback(account.AccountKey,
                LastBatch.Entries.Select(e => new ReviewEmailItem { SenderEmail = e.SenderEmail, Subject = e.Subject }).ToList(),
                $"Undo({LastBatch.Action})");

            LastBatch = null; // Clear after undo
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error undoing last batch");
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Updates pattern file based on user's feedback action.
    /// Returns (newPatternsAdded, patternsUpdated).
    /// </summary>
    private (int added, int updated) UpdatePatternsFromFeedback(string accountKey, List<ReviewEmailItem> items, string action)
    {
        var patterns = LoadPatterns(accountKey);
        int added = 0;
        int updated = 0;

        if (action == "delete")
        {
            // Strengthen or add sender rules
            var senderCounts = items
                .Where(i => !string.IsNullOrWhiteSpace(i.SenderEmail))
                .GroupBy(i => i.SenderEmail)
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var kvp in senderCounts)
            {
                var existing = patterns.FirstOrDefault(p =>
                    (p.Type == "Sender" || p.Type == "SenderWithUnsubscribe") &&
                    p.Value.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    // Boost confidence: move toward 0.95
                    existing.Weight = Math.Min(0.95, existing.Weight + 0.05);
                    existing.EvidenceCount += kvp.Value;
                    updated++;
                }
                else
                {
                    // Add new sender rule with high confidence (user explicitly deleted)
                    patterns.Add(new PatternRule
                    {
                        RuleId = $"rule_sender_{Guid.NewGuid().ToString("N")[..8]}",
                        Type = "Sender",
                        Value = kvp.Key,
                        Weight = 0.9,
                        EvidenceCount = kvp.Value
                    });
                    added++;
                }
            }
        }
        else if (action == "keep")
        {
            // Reduce confidence or remove rules for kept senders
            var keptSenders = new HashSet<string>(
                items.Where(i => !string.IsNullOrWhiteSpace(i.SenderEmail))
                     .Select(i => i.SenderEmail),
                StringComparer.OrdinalIgnoreCase);

            foreach (var sender in keptSenders)
            {
                var existing = patterns.FirstOrDefault(p =>
                    (p.Type == "Sender" || p.Type == "SenderWithUnsubscribe") &&
                    p.Value.Equals(sender, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    // Lower weight significantly — user explicitly wants to keep
                    existing.Weight = Math.Max(0.3, existing.Weight - 0.3);
                    updated++;
                }

                // Also check domain rules
                var domain = sender.Contains("@") ? sender.Split('@')[1] : "";
                if (!string.IsNullOrWhiteSpace(domain))
                {
                    var domainRule = patterns.FirstOrDefault(p =>
                        p.Type == "Domain" && p.Value.Equals(domain, StringComparison.OrdinalIgnoreCase));

                    if (domainRule != null)
                    {
                        domainRule.Weight = Math.Max(0.3, domainRule.Weight - 0.2);
                        updated++;
                    }
                }
            }
        }

        SavePatternFile(accountKey, patterns);
        return (added, updated);
    }

    /// <summary>
    /// Appends feedback entries to feedback_log.csv.
    /// </summary>
    private void LogFeedback(string accountKey, List<ReviewEmailItem> items, string action)
    {
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EmailManage", "accounts", accountKey);
            Directory.CreateDirectory(dir);

            var path = Path.Combine(dir, "feedback_log.csv");
            bool isNew = !File.Exists(path);

            using var writer = new StreamWriter(path, append: true);
            if (isNew)
            {
                writer.WriteLine("timestamp,action,sender_email,domain,subject,entry_id,matched_rules");
            }

            var timestamp = DateTime.UtcNow.ToString("O");
            foreach (var item in items)
            {
                var escapedSubject = EscapeCsv(item.Subject);
                var rules = EscapeCsv(item.MatchedRulesSummary);
                writer.WriteLine($"{timestamp},{action},{item.SenderEmail},{item.Domain},{escapedSubject},{item.EntryId},{rules}");
            }

            _log.Info("Logged {Count} feedback entries for action '{Action}'", items.Count, action);
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to write feedback log");
        }
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    private List<PatternRule> LoadPatterns(string accountKey)
    {
        var patterns = new List<PatternRule>();
        var path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey, "pattern_v1.yaml");
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
                    currentRule = new PatternRule { RuleId = trimmed[5..].Trim() };
                }
                else if (currentRule != null)
                {
                    if (trimmed.StartsWith("type:")) currentRule.Type = trimmed[5..].Trim();
                    else if (trimmed.StartsWith("value:")) currentRule.Value = trimmed[6..].Trim().Trim('"');
                    else if (trimmed.StartsWith("weight:"))
                    {
                        if (double.TryParse(trimmed[7..].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double w))
                            currentRule.Weight = w;
                    }
                    else if (trimmed.StartsWith("evidence_count:"))
                    {
                        if (int.TryParse(trimmed[15..].Trim(), out int e))
                            currentRule.EvidenceCount = e;
                    }
                }
            }
            if (currentRule != null) patterns.Add(currentRule);
        }
        catch (Exception ex) { _log.Error(ex, "Failed to load pattern file"); }

        return patterns;
    }

    private void SavePatternFile(string accountKey, List<PatternRule> patterns)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey);
        Directory.CreateDirectory(dir);

        var path = Path.Combine(dir, "pattern_v1.yaml");
        using var writer = new StreamWriter(path);
        writer.WriteLine("version: 1");
        writer.WriteLine($"last_updated: {DateTime.UtcNow:O}");
        writer.WriteLine("rules:");
        foreach (var p in patterns)
        {
            writer.WriteLine($"  - id: {p.RuleId}");
            writer.WriteLine($"    type: {p.Type}");
            writer.WriteLine($"    value: \"{p.Value}\"");
            writer.WriteLine($"    weight: {p.Weight.ToString("F2", CultureInfo.InvariantCulture)}");
            writer.WriteLine($"    evidence_count: {p.EvidenceCount}");
        }
    }

    private EmailScore ScoreEmail(EmailMessageInfo email, List<PatternRule> patterns)
    {
        var score = new EmailScore();
        string sender = email.SenderEmailAddress?.ToLowerInvariant() ?? "";
        string domain = sender.Contains("@") ? sender.Split('@')[1] : "";
        string subject = email.Subject?.Trim() ?? "";
        string body = email.Body ?? "";

        foreach (var rule in patterns)
        {
            bool matched = rule.Type switch
            {
                "Sender" => sender == rule.Value.ToLowerInvariant(),
                "Domain" => domain == rule.Value.ToLowerInvariant(),
                "Subject" => subject == rule.Value,
                "SenderWithUnsubscribe" => sender == rule.Value.ToLowerInvariant() &&
                                           body.Contains("unsubscribe", StringComparison.OrdinalIgnoreCase),
                _ => false
            };

            if (matched)
            {
                score.MatchedRules.Add(rule);
                if (rule.Weight > score.Confidence)
                    score.Confidence = rule.Weight;
            }
        }

        return score;
    }
}
