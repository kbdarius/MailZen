namespace EmailManage.Models;

/// <summary>
/// Represents metadata about a single Outlook email account/store.
/// </summary>
public sealed class OutlookAccountInfo
{
    /// <summary>Display name for the account (e.g., "Keivan - ZodVest").</summary>
    public string DisplayName { get; init; } = string.Empty;

    /// <summary>Primary email address (e.g., "keivan@zodvest.com").</summary>
    public string EmailAddress { get; init; } = string.Empty;

    /// <summary>Outlook store name as reported by the profile.</summary>
    public string StoreName { get; init; } = string.Empty;

    /// <summary>File path to the store (.ost / .pst) if available.</summary>
    public string StoreFilePath { get; init; } = string.Empty;

    /// <summary>Provider hint: "Exchange", "IMAP", "POP3", "Outlook.com", "Unknown".</summary>
    public string ProviderHint { get; init; } = "Unknown";

    /// <summary>Whether the store is currently connected and accessible.</summary>
    public bool IsConnected { get; set; }

    /// <summary>Sanitized key for file-system storage (e.g., "keivan_at_zodvest_com").</summary>
    public string AccountKey => SanitizeKey(EmailAddress);

    /// <summary>Number of items in the Inbox folder (populated on demand).</summary>
    public int InboxCount { get; set; }

    /// <summary>Number of items in Deleted Items folder (populated on demand).</summary>
    public int DeletedItemsCount { get; set; }

    /// <summary>Whether a pattern file already exists for this account.</summary>
    public bool HasPatternFile { get; set; }

    private static string SanitizeKey(string email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return "unknown";

        return email
            .ToLowerInvariant()
            .Replace("@", "_at_")
            .Replace(".", "_")
            .Replace(" ", "_");
    }
}
