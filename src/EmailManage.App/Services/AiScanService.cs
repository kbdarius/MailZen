using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Orchestrates an AI scan of remaining inbox emails, comparing AI classification
/// against rule-based results to measure the gap.
/// </summary>
public class AiScanService
{
    private readonly OutlookConnectorService _connector;
    private readonly OllamaClient _aiClient;
    private readonly DiagnosticLogger _log;

    public AiScanService(OutlookConnectorService connector, string modelName = "gemma3:4b")
    {
        _connector = connector;
        _aiClient = new OllamaClient(modelName);
        _log = DiagnosticLogger.Instance;
    }

    /// <summary>
    /// Scans inbox emails with AI classification and compares against rule-based scores.
    /// Returns a benchmark report.
    /// </summary>
    public async Task<AiScanResult> RunBenchmarkAsync(
        OutlookAccountInfo account, int maxEmails = 100,
        IProgress<string>? progress = null, CancellationToken ct = default)
    {
        var result = new AiScanResult();
        var sw = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            // 1. Load inbox emails (ones that triage did NOT catch)
            progress?.Report("Fetching inbox emails...");
            var inboxEmails = await _connector.GetEmailsAsync(
                account.EmailAddress, account.StoreName,
                OutlookFolderType.Inbox, 1, maxEmails, skipUnread: false);

            result.TotalScanned = inboxEmails.Count;
            progress?.Report($"Found {inboxEmails.Count} inbox emails. Starting AI scan...");

            // 2. Load existing patterns for rule comparison
            var patterns = LoadPatterns(account.AccountKey);

            // 3. Classify each email with AI
            int processed = 0;
            foreach (var email in inboxEmails)
            {
                ct.ThrowIfCancellationRequested();

                processed++;
                progress?.Report($"AI scanning {processed}/{inboxEmails.Count}: {email.SenderName}...");

                bool hasUnsubscribe = (email.Body ?? "").Contains("unsubscribe", StringComparison.OrdinalIgnoreCase);
                string bodyPreview = (email.Body ?? "").Length > 200 ? email.Body![..200] : (email.Body ?? "");

                var aiResult = await _aiClient.ClassifyEmailAsync(
                    email.SenderEmailAddress, email.Subject, bodyPreview, hasUnsubscribe, ct);

                // Get rule-based score
                var ruleScore = ScoreEmail(email, patterns);

                var item = new AiScanItem
                {
                    EntryId = email.EntryId,
                    Subject = email.Subject,
                    SenderName = email.SenderName,
                    SenderEmail = email.SenderEmailAddress?.ToLowerInvariant() ?? "",
                    Domain = ExtractDomain(email.SenderEmailAddress),
                    ReceivedTime = email.ReceivedTime,

                    AiClassification = aiResult.Classification,
                    AiConfidence = aiResult.Confidence,
                    AiReason = aiResult.Reason,
                    AiLatencyMs = aiResult.LatencyMs,

                    RuleConfidence = ruleScore.Confidence,
                    RuleMatched = ruleScore.MatchedRules.Count > 0,
                    RuleMatchSummary = string.Join(", ", ruleScore.MatchedRules.Select(r => $"{r.Type}:{r.Value}")),

                    // Key metric: AI says junk but rules didn't catch it
                    IsAiGap = aiResult.IsJunk && ruleScore.Confidence < 0.7
                };

                result.Items.Add(item);
            }

            // 4. Compute summary stats
            result.AiFlaggedJunk = result.Items.Count(i => i.AiClassification == "JUNK");
            result.RulesFlaggedJunk = result.Items.Count(i => i.RuleConfidence >= 0.7);
            result.GapCount = result.Items.Count(i => i.IsAiGap);
            result.AvgLatencyMs = result.Items.Count > 0 ? (long)result.Items.Average(i => i.AiLatencyMs) : 0;
            result.TotalTimeSeconds = (int)sw.Elapsed.TotalSeconds;
            result.Success = true;

            progress?.Report(
                $"AI scan complete! {result.AiFlaggedJunk} flagged as junk. " +
                $"Gap: {result.GapCount} emails AI caught that rules missed. " +
                $"({result.TotalTimeSeconds}s total, ~{result.AvgLatencyMs}ms/email)");
        }
        catch (OperationCanceledException)
        {
            result.ErrorMessage = "Scan cancelled.";
            progress?.Report("Scan cancelled.");
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error during AI benchmark scan");
            result.ErrorMessage = ex.Message;
            progress?.Report($"Error: {ex.Message}");
        }

        return result;
    }

    private static string ExtractDomain(string? email)
    {
        if (string.IsNullOrWhiteSpace(email) || !email.Contains('@')) return "";
        return email.ToLowerInvariant().Split('@')[1];
    }

    private List<PatternRule> LoadPatterns(string accountKey)
    {
        var patterns = new List<PatternRule>();
        var path = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey, "pattern_v1.yaml");
        if (!System.IO.File.Exists(path)) return patterns;

        try
        {
            var lines = System.IO.File.ReadAllLines(path);
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
                        if (double.TryParse(trimmed[7..].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double w))
                            currentRule.Weight = w;
                    }
                    else if (trimmed.StartsWith("evidence_count:"))
                    {
                        if (int.TryParse(trimmed[15..].Trim(), out int e)) currentRule.EvidenceCount = e;
                    }
                }
            }
            if (currentRule != null) patterns.Add(currentRule);
        }
        catch { }

        return patterns;
    }

    private static EmailScore ScoreEmail(EmailMessageInfo email, List<PatternRule> patterns)
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

/// <summary>
/// Result of an AI benchmark scan.
/// </summary>
public class AiScanResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int TotalScanned { get; set; }
    public int AiFlaggedJunk { get; set; }
    public int RulesFlaggedJunk { get; set; }
    public int GapCount { get; set; } // AI caught but rules missed
    public long AvgLatencyMs { get; set; }
    public int TotalTimeSeconds { get; set; }
    public List<AiScanItem> Items { get; set; } = new();
}

/// <summary>
/// One email's AI scan result with rule comparison.
/// </summary>
public class AiScanItem
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public string Domain { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }

    // AI classification
    public string AiClassification { get; set; } = string.Empty;
    public double AiConfidence { get; set; }
    public string AiReason { get; set; } = string.Empty;
    public long AiLatencyMs { get; set; }

    // Rule-based comparison
    public double RuleConfidence { get; set; }
    public bool RuleMatched { get; set; }
    public string RuleMatchSummary { get; set; } = string.Empty;

    // Gap analysis
    public bool IsAiGap { get; set; } // AI says junk, rules say keep

    public bool IsSelected { get; set; }
}
