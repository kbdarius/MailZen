using EmailManage.Models;
namespace EmailManage.Services;

public sealed class SearchOrchestrator : ISearchOrchestrator
{
    private readonly ILocalSearchService _localSearch;
    private readonly IAiSearchProvider? _aiSearch;
    public SearchOrchestrator(ILocalSearchService localSearch, IAiSearchProvider? aiSearch = null)
    { _localSearch = localSearch; _aiSearch = aiSearch; }

    public async Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, bool useAi, CancellationToken cancellationToken = default)
    {
        var boundedRequest = request with { MaxResults = Math.Clamp(request.MaxResults, 1, 100) };
        if (!useAi || _aiSearch is null)
            return await _localSearch.SearchAsync(boundedRequest, cancellationToken);

        var intent = await _aiSearch.ParseIntentAsync(request.Query, request.Scope, cancellationToken);
        var intentQuery = SearchQueryText.BuildIntentQuery(intent);
        var retrievalRequest = string.IsNullOrWhiteSpace(intentQuery)
            ? boundedRequest
            : boundedRequest with { Query = intentQuery, Mode = SearchMode.SmartLocal };
        var candidates = await _localSearch.SearchAsync(retrievalRequest, cancellationToken);
        if (candidates.Count == 0) return candidates;
        var boundedCandidates = candidates.Take(20).ToArray();
        var ranked = await _aiSearch.RerankAsync(boundedCandidates, intent, cancellationToken);
        var byId = candidates.ToDictionary(c => c.Message.MessageId, StringComparer.Ordinal);
        return ranked.Where(r => byId.ContainsKey(r.MessageId)).Select(r => byId[r.MessageId] with { Score = r.Score, Explanation = r.Explanation }).ToArray();
    }
}
