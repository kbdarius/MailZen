namespace EmailManage.Models;

public sealed record AccountScope(string AccountId, string DisplayName, string EmailAddress);

public sealed record SearchScope(
    IReadOnlySet<string> AccountIds,
    IReadOnlySet<string>? FolderIds = null,
    DateTime? ReceivedAfterUtc = null,
    DateTime? ReceivedBeforeUtc = null,
    bool? IsUnread = null,
    bool? HasAttachments = null,
    bool IncludeSentItems = false);

public enum SearchMode { ConversationalAi, SmartLocal, Boolean }

public sealed record SearchRequest(string Query, SearchScope Scope, int MaxResults = 50, SearchMode Mode = SearchMode.SmartLocal);

public sealed record SearchIntent
{
    public IReadOnlyList<string> People { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Organizations { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> RequiredKeywords { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> OptionalKeywords { get; init; } = Array.Empty<string>();
    public DateTime? ReceivedAfterUtc { get; init; }
    public DateTime? ReceivedBeforeUtc { get; init; }
    public string? SortPreference { get; init; }
    public string? AmbiguityNote { get; init; }
}

public sealed record IndexedMessage(
    string MessageId,
    string AccountId,
    string FolderId,
    string StoreId,
    string EntryId,
    string? InternetMessageId,
    string Subject,
    string SenderName,
    string SenderAddress,
    string BodyText,
    DateTime ReceivedUtc,
    bool IsUnread,
    bool HasAttachments,
    string AttachmentNames,
    string? ConversationId = null,
    string FolderType = "Inbox");

public sealed record SearchResult(IndexedMessage Message, double Score, string Excerpt, string? Explanation = null);

public sealed record IndexCoverage(int MessageCount, DateTime? EarliestReceivedUtc, DateTime? LatestReceivedUtc);

public sealed record RankedCandidate(string MessageId, double Score, string? Explanation);

public sealed record OutlookAccount(string AccountId, string StoreId, string DisplayName, string EmailAddress);

public sealed record OutlookFolder(string FolderId, string AccountId, string EntryId, string Path, string FolderType);

public sealed record OutlookReadOptions(DateTime? ReceivedAfterUtc, DateTime? ReceivedBeforeUtc, int BatchSize = 250);
