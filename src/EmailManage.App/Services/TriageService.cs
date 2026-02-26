using System.IO;
using System.Text.Json;
using EmailManage.Models;

namespace EmailManage.Services;

public class TriageService
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;

    public TriageService(OutlookConnectorService connector)
    {
        _connector = connector;
        _log = DiagnosticLogger.Instance;
    }

    public async Task<TriageResult> RunTriageAsync(OutlookAccountInfo account, IProgress<string>? progress = null)
    {
        var result = new TriageResult();
        try
        {
            progress?.Report("Loading patterns...");
            _log.Info("Triage: Loading patterns for account key [{AccountKey}]", account.AccountKey);
            var patterns = LoadPatterns(account.AccountKey);
            _log.Info("Triage: Loaded {Count} patterns", patterns.Count);
            if (patterns.Count == 0)
            {
                result.Success = false;
                result.ErrorMessage = "No patterns found. Please run 'Analyze Patterns' first.";
                return result;
            }

            progress?.Report("Ensuring triage folders exist...");
            bool foldersCreated = await _connector.EnsureTriageFoldersAsync(account.EmailAddress, account.StoreName);
            if (!foldersCreated)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to create or access Smart Cleanup folders in Outlook.";
                return result;
            }

            progress?.Report("Fetching recent inbox emails...");
            _log.Info("Triage: Fetching emails for email=[{Email}] store=[{Store}]", account.EmailAddress, account.StoreName);
            // Fetch last 1 month of inbox emails for triage (max 500 to be safe)
            var inboxEmails = await _connector.GetEmailsAsync(account.EmailAddress, account.StoreName, OutlookFolderType.Inbox, 1, 500, skipUnread: false);
            _log.Info("Triage: Fetched {Count} inbox emails", inboxEmails.Count);
            
            progress?.Report($"Scoring {inboxEmails.Count} emails...");
            
            int movedToDelete = 0;
            int movedToReview = 0;

            foreach (var email in inboxEmails)
            {
                var score = ScoreEmail(email, patterns);
                
                if (score.Confidence >= 0.9)
                {
                    // High confidence -> Delete Candidates
                    bool moved = await _connector.MoveEmailAsync(email.EntryId, account.EmailAddress, account.StoreName, OutlookConnectorService.DeleteCandidatesFolderName);
                    if (moved) movedToDelete++;
                }
                else if (score.Confidence >= 0.7)
                {
                    // Medium confidence -> Needs Review
                    bool moved = await _connector.MoveEmailAsync(email.EntryId, account.EmailAddress, account.StoreName, OutlookConnectorService.NeedsReviewFolderName);
                    if (moved) movedToReview++;
                }
            }

            result.Success = true;
            result.ScannedCount = inboxEmails.Count;
            result.MovedToDeleteCandidates = movedToDelete;
            result.MovedToNeedsReview = movedToReview;

            progress?.Report($"Triage complete. Moved {movedToDelete} to Delete Candidates, {movedToReview} to Needs Review.");
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Error running triage for account {Account}", account.EmailAddress);
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    private List<PatternRule> LoadPatterns(string accountKey)
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
            _log.Error(ex, "Failed to parse pattern file");
        }

        return patterns;
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
            bool matched = false;
            switch (rule.Type)
            {
                case "Sender":
                    matched = sender == rule.Value.ToLowerInvariant();
                    break;
                case "Domain":
                    matched = domain == rule.Value.ToLowerInvariant();
                    break;
                case "Subject":
                    matched = subject == rule.Value;
                    break;
                case "SenderWithUnsubscribe":
                    matched = sender == rule.Value.ToLowerInvariant() && body.Contains("unsubscribe", StringComparison.OrdinalIgnoreCase);
                    break;
            }

            if (matched)
            {
                score.MatchedRules.Add(rule);
                // Simple max confidence for now
                if (rule.Weight > score.Confidence)
                {
                    score.Confidence = rule.Weight;
                }
            }
        }

        return score;
    }
}

public class TriageResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int ScannedCount { get; set; }
    public int MovedToDeleteCandidates { get; set; }
    public int MovedToNeedsReview { get; set; }
}

public class EmailScore
{
    public double Confidence { get; set; }
    public List<PatternRule> MatchedRules { get; set; } = new();
}