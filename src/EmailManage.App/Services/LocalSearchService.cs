using EmailManage.Models;

namespace EmailManage.Services;

public sealed class LocalSearchService : ILocalSearchService
{
    private readonly MailZenDatabase _database;
    public LocalSearchService(MailZenDatabase database) => _database = database;
    public Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, CancellationToken cancellationToken = default) =>
        _database.SearchAsync(request, cancellationToken);
}
