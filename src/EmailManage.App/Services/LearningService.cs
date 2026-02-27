using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Scans the user's Deleted Items folder to build a <see cref="LearnedProfile"/>
/// of senders and domains they typically delete.
/// This profile is later injected into Ollama prompts during triage.
/// </summary>
public class LearningService
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;

    public LearningService(OutlookConnectorService connector)
    {
        _connector = connector;
        _log = DiagnosticLogger.Instance;
    }

    /// <summary>
    /// Scans Deleted Items for the given account, builds sender/domain frequency
    /// tables, saves the profile locally, and returns it.
    /// </summary>
    public async Task<LearnedProfile> LearnFromDeletedItemsAsync(
        OutlookAccountInfo account,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var profile = new LearnedProfile
        {
            AccountKey = account.AccountKey,
            LearnedAt = DateTime.UtcNow
        };

        progress?.Report("Fetching deleted emails (last 12 months)...");
        _log.Info("Learning: scanning Deleted Items for {Account}", account.EmailAddress);

        // Fetch deleted items — read emails from the last 12 months, max 500
        var deletedEmails = await _connector.GetEmailsAsync(
            account.EmailAddress,
            account.StoreName,
            OutlookConnectorService.OlFolderDeletedItems,
            maxMonths: 12,
            maxItems: 500,
            skipUnread: false); // We want ALL deleted items, read or unread

        ct.ThrowIfCancellationRequested();

        profile.TotalDeletedScanned = deletedEmails.Count;
        _log.Info("Learning: found {Count} deleted emails to analyse", deletedEmails.Count);

        if (deletedEmails.Count == 0)
        {
            progress?.Report("No deleted emails found. Using general AI knowledge only.");
            profile.Save();
            return profile;
        }

        // Build frequency tables
        progress?.Report($"Analysing {deletedEmails.Count} deleted emails...");

        foreach (var email in deletedEmails)
        {
            ct.ThrowIfCancellationRequested();

            var sender = (email.SenderEmailAddress ?? "").Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(sender)) continue;

            // Sender frequency
            if (profile.DeletedSenderCounts.ContainsKey(sender))
                profile.DeletedSenderCounts[sender]++;
            else
                profile.DeletedSenderCounts[sender] = 1;

            // Domain frequency
            var atIndex = sender.IndexOf('@');
            if (atIndex > 0 && atIndex < sender.Length - 1)
            {
                var domain = sender[(atIndex + 1)..];
                if (profile.DeletedDomainCounts.ContainsKey(domain))
                    profile.DeletedDomainCounts[domain]++;
                else
                    profile.DeletedDomainCounts[domain] = 1;
            }
        }

        // Persist
        profile.Save();

        var topDomains = profile.DeletedDomainCounts
            .OrderByDescending(kv => kv.Value)
            .Take(5)
            .Select(kv => $"{kv.Key} ({kv.Value}x)");

        progress?.Report(
            $"Learned from {deletedEmails.Count} emails. " +
            $"Top domains: {string.Join(", ", topDomains)}");

        _log.Info("Learning complete: {Senders} unique senders, {Domains} unique domains",
            profile.DeletedSenderCounts.Count, profile.DeletedDomainCounts.Count);

        return profile;
    }
}
