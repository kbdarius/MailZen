using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Result of a triage run — how many emails were scanned, flagged, and moved.
/// </summary>
public class TriageResult
{
    public int TotalScanned { get; set; }
    public int FlaggedAsJunk { get; set; }
    public int MovedToReview { get; set; }
    public int Errors { get; set; }
    public List<AiClassification> Classifications { get; set; } = new();
}

/// <summary>
/// Scans the user's Inbox, classifies each email via Ollama,
/// and moves junk to the "Review for Deletion" folder.
/// Uses the <see cref="LearnedProfile"/> for personalised context.
/// </summary>
public class TriageService
{
    private readonly OutlookConnectorService _connector;
    private readonly OllamaClient _ollama;
    private readonly DiagnosticLogger _log;

    public TriageService(OutlookConnectorService connector, OllamaClient ollama)
    {
        _connector = connector;
        _ollama = ollama;
        _log = DiagnosticLogger.Instance;
    }

    /// <summary>
    /// Runs a full triage: fetch inbox emails → classify each with Ollama →
    /// bulk-move JUNK to "Review for Deletion".
    /// Returns a <see cref="TriageResult"/> with counts and details.
    /// </summary>
    public async Task<TriageResult> TriageInboxAsync(
        OutlookAccountInfo account,
        LearnedProfile profile,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var result = new TriageResult();

        // 1. Ensure the review folder exists
        progress?.Report("Preparing Review for Deletion folder...");
        await _connector.EnsureReviewFolderAsync(account.EmailAddress, account.StoreName);

        // 2. Fetch inbox emails (read-only, last 6 months, max 200)
        progress?.Report("Fetching inbox emails...");
        var inboxEmails = await _connector.GetEmailsAsync(
            account.EmailAddress,
            account.StoreName,
            OutlookConnectorService.OlFolderInbox,
            maxMonths: 6,
            maxItems: 200,
            skipUnread: false); // Scan all emails, unread and read

        ct.ThrowIfCancellationRequested();

        result.TotalScanned = inboxEmails.Count;
        _log.Info("Triage: fetched {Count} inbox emails for classification", inboxEmails.Count);

        if (inboxEmails.Count == 0)
        {
            progress?.Report("Inbox is empty — nothing to triage.");
            return result;
        }

        // 3. Build the profile context for AI prompts
        var profileContext = profile.ToPromptContext();

        // 4. Classify each email with Ollama
        var junkEntryIds = new List<string>();
        bool wasCancelled = false;

        for (int i = 0; i < inboxEmails.Count; i++)
        {
            if (ct.IsCancellationRequested)
            {
                wasCancelled = true;
                _log.Info("Triage: cancelled at email {Index}/{Total}", i, inboxEmails.Count);
                break;
            }

            var email = inboxEmails[i];

            // Skip senders the user explicitly protected (kept in previous review)
            var senderLower = (email.SenderEmailAddress ?? "").ToLowerInvariant();
            if (!string.IsNullOrEmpty(senderLower) && profile.DoNotDeleteSenders.Contains(senderLower))
            {
                _log.Debug("SKIP (protected sender): [{Sender}]", senderLower);
                continue;
            }

            // Skip senders that already have Outlook rules — they'll be auto-deleted going forward
            if (!string.IsNullOrEmpty(senderLower) && profile.RuleCreatedSenders.Contains(senderLower))
            {
                _log.Debug("SKIP (has rule): [{Sender}]", senderLower);
                continue;
            }

            var subjectPreview = email.Subject?.Length > 50
                ? email.Subject[..50] + "..."
                : email.Subject ?? "(no subject)";

            progress?.Report($"Classifying {i + 1}/{inboxEmails.Count}: {subjectPreview}");

            try
            {
                // Detect unsubscribe link in body
                var bodyLower = (email.Body ?? "").ToLowerInvariant();
                bool hasUnsubscribe = bodyLower.Contains("unsubscribe")
                    || bodyLower.Contains("opt out")
                    || bodyLower.Contains("opt-out");

                // Build body preview (first 200 chars)
                var bodyPreview = (email.Body ?? "").Length > 200
                    ? email.Body![..200]
                    : email.Body ?? "";

                // If we have a learned profile, prepend context to body preview
                // so Ollama sees the user's deletion history
                var enrichedBody = string.IsNullOrWhiteSpace(profileContext)
                    ? bodyPreview
                    : profileContext + "\n---\nEMAIL BODY PREVIEW:\n" + bodyPreview;

                var classification = await _ollama.ClassifyEmailAsync(
                    email.SenderEmailAddress ?? "",
                    email.Subject ?? "",
                    enrichedBody,
                    hasUnsubscribe,
                    ct);

                result.Classifications.Add(classification);

                if (classification.IsJunk)
                {
                    junkEntryIds.Add(email.EntryId);
                    _log.Debug("JUNK: [{Sender}] {Subject} (conf={Conf:P0})",
                        email.SenderEmailAddress, subjectPreview, classification.Confidence);
                }
            }
            catch (Exception ex)
            {
                result.Errors++;
                _log.Warn("Triage: error classifying email {Index}: {Error}", i, ex.Message);
            }
        }

        result.FlaggedAsJunk = junkEntryIds.Count;
        _log.Info("Triage: {Junk}/{Total} classified as JUNK, {Errors} errors",
            junkEntryIds.Count, inboxEmails.Count, result.Errors);

        // 5. Bulk-move junk to "Review for Deletion" (safe — even if cancelled, move what we have)
        if (junkEntryIds.Count > 0)
        {
            progress?.Report($"Moving {junkEntryIds.Count} emails to Review for Deletion...");
            result.MovedToReview = await _connector.BulkMoveToReviewFolderAsync(
                junkEntryIds, account.EmailAddress, account.StoreName, progress);

            _log.Info("Triage: moved {Moved} emails to Review for Deletion", result.MovedToReview);
        }
        else if (!wasCancelled)
        {
            progress?.Report("No junk detected — your inbox looks clean!");
        }

        // After moving partial results safely, propagate cancellation
        if (wasCancelled)
        {
            progress?.Report($"Stopped. Moved {result.MovedToReview} emails found so far.");
            throw new OperationCanceledException(ct);
        }

        return result;
    }
}
