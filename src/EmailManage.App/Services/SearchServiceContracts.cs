using EmailManage.Models;

namespace EmailManage.Services;

public interface IOutlookReadService
{
    Task<IReadOnlyList<OutlookAccount>> GetAccountsAsync(CancellationToken cancellationToken = default);
    IAsyncEnumerable<OutlookFolder> EnumerateFoldersAsync(string accountId, CancellationToken cancellationToken = default);
    IAsyncEnumerable<IndexedMessage> ReadMessagesAsync(OutlookFolder folder, OutlookReadOptions options, CancellationToken cancellationToken = default);
}

public interface IOutlookOpenService
{
    Task<bool> TryOpenAsync(IndexedMessage message, CancellationToken cancellationToken = default);
    Task<string> ExportToMsgAsync(IndexedMessage message, CancellationToken cancellationToken = default);
}

public interface IEmailIndexService
{
    Task SyncAsync(IReadOnlySet<string> accountIds, DateTime? sinceUtc, DateTime? beforeUtc = null, IProgress<string>? progress = null, CancellationToken cancellationToken = default);
}

public interface ILocalSearchService
{
    Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, CancellationToken cancellationToken = default);
}

public interface IAiSearchProvider
{
    Task<SearchIntent> ParseIntentAsync(string query, SearchScope scope, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<RankedCandidate>> RerankAsync(IReadOnlyList<SearchResult> candidates, SearchIntent intent, CancellationToken cancellationToken = default);
    Task<bool> CheckHealthAsync(CancellationToken cancellationToken = default);
}

public interface ISearchOrchestrator
{
    Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, bool useAi, CancellationToken cancellationToken = default);
}

public interface ICredentialStore
{
    Task SetApiKeyAsync(string apiKey, CancellationToken cancellationToken = default);
    Task<bool> HasApiKeyAsync(CancellationToken cancellationToken = default);
    Task RemoveApiKeyAsync(CancellationToken cancellationToken = default);
}

public interface ISyncScheduler
{
    Task ConfigureDailyAsync(TimeOnly localTime, CancellationToken cancellationToken = default);
    Task RemoveAsync(CancellationToken cancellationToken = default);
    Task<bool> ExistsAsync(CancellationToken cancellationToken = default);
}
