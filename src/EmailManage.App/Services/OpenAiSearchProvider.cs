using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.IO;
using EmailManage.Models;

namespace EmailManage.Services;

public enum SearchModelProfile { Fast, Smart }

public sealed record SearchModelDefinition(SearchModelProfile Profile, string DisplayName, string ModelId, string CostNote);

public static class MailZenModelRegistry
{
    public const string ApiKeyEnvironmentVariable = "OPENAI_API_KEY";
    public const string ModelEnvironmentVariable = "MAILZEN_OPENAI_MODEL";
    public const string FastModelId = "gpt-4o-mini";
    public const string NanoModelId = "gpt-5-nano";
    public const string SmartModelId = "gpt-5.6-luna";
    public static IReadOnlyList<string> SupportedModelIds { get; } = new[] { NanoModelId, SmartModelId, FastModelId };
    public static IReadOnlyList<SearchModelDefinition> Profiles { get; } = new[]
    {
        new SearchModelDefinition(SearchModelProfile.Fast, "Fast / Economy", FastModelId, "Lower cost and latency"),
        new SearchModelDefinition(SearchModelProfile.Smart, "Smart", SmartModelId, "Higher reasoning cost")
    };

    public static SearchModelDefinition Resolve(SearchModelProfile profile)
    {
        var configured = Environment.GetEnvironmentVariable(ModelEnvironmentVariable, EnvironmentVariableTarget.Process)
            ?? Environment.GetEnvironmentVariable(ModelEnvironmentVariable, EnvironmentVariableTarget.User);
        return Profiles.FirstOrDefault(p => string.Equals(p.ModelId, configured, StringComparison.OrdinalIgnoreCase))
            ?? Profiles.First(p => p.Profile == profile);
    }

    public static string ResolveModelId(string? modelId, SearchModelProfile fallback = SearchModelProfile.Fast) =>
        SupportedModelIds.FirstOrDefault(id => string.Equals(id, modelId, StringComparison.OrdinalIgnoreCase))
        ?? Resolve(fallback).ModelId;
}

public sealed class OpenAiSearchProvider : IAiSearchProvider
{
    private readonly HttpClient _httpClient;
    private readonly SearchModelProfile _profile;
    private readonly string? _modelId;
    private readonly OpenAiCredentialStore _credentialStore;
    private const string Endpoint = "https://api.openai.com/v1/responses";

    public OpenAiSearchProvider(HttpClient httpClient, SearchModelProfile profile = SearchModelProfile.Fast, string? modelId = null)
    { _httpClient = httpClient; _profile = profile; _modelId = MailZenModelRegistry.ResolveModelId(modelId, profile); _credentialStore = new OpenAiCredentialStore(); }

    public async Task<SearchIntent> ParseIntentAsync(string query, SearchScope scope, CancellationToken cancellationToken = default)
    {
        var json = await SendStructuredAsync(
            "Interpret the email search request. Never add accounts or filters beyond the user-selected scope. Return only the schema.",
            $"Selected account IDs: {string.Join(", ", scope.AccountIds)}\nRequest: {query}", IntentSchema, cancellationToken);
        var intent = JsonSerializer.Deserialize<SearchIntent>(json, JsonOptions) ?? throw new InvalidDataException("OpenAI returned an empty search intent.");
        return intent with
        {
            ReceivedAfterUtc = scope.ReceivedAfterUtc ?? intent.ReceivedAfterUtc,
            ReceivedBeforeUtc = scope.ReceivedBeforeUtc ?? intent.ReceivedBeforeUtc
        };
    }

    public async Task<IReadOnlyList<RankedCandidate>> RerankAsync(IReadOnlyList<SearchResult> candidates, SearchIntent intent, CancellationToken cancellationToken = default)
    {
        var bounded = candidates.Take(20).Select(c => new { id = c.Message.MessageId, subject = c.Message.Subject, sender = c.Message.SenderName, excerpt = c.Excerpt }).ToArray();
        var json = await SendStructuredAsync("Rank only the supplied candidate IDs. Do not invent IDs. Explain the match briefly.",
            JsonSerializer.Serialize(new { intent, candidates = bounded }, JsonOptions), RankingSchema, cancellationToken);
        var response = JsonSerializer.Deserialize<RankingResponse>(json, JsonOptions) ?? new RankingResponse();
        var allowed = candidates.Select(c => c.Message.MessageId).ToHashSet(StringComparer.Ordinal);
        return response.Candidates.Where(c => allowed.Contains(c.MessageId)).Take(20).ToArray();
    }

    public async Task<bool> CheckHealthAsync(CancellationToken cancellationToken = default)
    {
        var key = GetApiKey();
        if (string.IsNullOrWhiteSpace(key)) return false;
        using var request = CreateRequest(key);
        request.Content = new StringContent(JsonSerializer.Serialize(new { model = ModelId, input = "Reply with OK.", max_output_tokens = 32 }, JsonOptions), Encoding.UTF8, "application/json");
        try { using var response = await _httpClient.SendAsync(request, cancellationToken); return response.IsSuccessStatusCode; }
        catch (HttpRequestException) { return false; }
    }

    private async Task<string> SendStructuredAsync(string instructions, string input, object schema, CancellationToken cancellationToken)
    {
        var key = GetApiKey() ?? throw new InvalidOperationException("OpenAI API key is not configured.");
        var model = ModelId;
        var outputTokenBudget = schema == RankingSchema ? 4000 : 1200;
        var payload = new Dictionary<string, object?>
        {
            ["model"] = model, ["instructions"] = instructions, ["input"] = input, ["max_output_tokens"] = outputTokenBudget,
            ["text"] = new { format = new { type = "json_schema", name = "mailzen_search", strict = true, schema } }
        };
        if (model.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase)) payload["text"] = new { verbosity = "low", format = new { type = "json_schema", name = "mailzen_search", strict = true, schema } };
        using var request = CreateRequest(key);
        request.Content = new StringContent(JsonSerializer.Serialize(payload, JsonOptions), Encoding.UTF8, "application/json");
        using var response = await _httpClient.SendAsync(request, cancellationToken);
        var responseText = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode) throw new HttpRequestException($"OpenAI returned HTTP {(int)response.StatusCode}.");
        using var document = JsonDocument.Parse(responseText);
        foreach (var item in document.RootElement.GetProperty("output").EnumerateArray())
            foreach (var part in item.GetProperty("content").EnumerateArray())
                if (part.GetProperty("type").GetString() == "output_text") return part.GetProperty("text").GetString() ?? throw new InvalidDataException("Missing structured output.");
        throw new InvalidDataException("OpenAI response did not contain output text.");
    }

    private HttpRequestMessage CreateRequest(string key)
    { var request = new HttpRequestMessage(HttpMethod.Post, Endpoint); request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key); return request; }
    private string? GetApiKey() => Environment.GetEnvironmentVariable(MailZenModelRegistry.ApiKeyEnvironmentVariable, EnvironmentVariableTarget.Process)
        ?? Environment.GetEnvironmentVariable(MailZenModelRegistry.ApiKeyEnvironmentVariable, EnvironmentVariableTarget.User)
        ?? _credentialStore.TryReadApiKey();
    private string ModelId => _modelId ?? MailZenModelRegistry.Resolve(_profile).ModelId;
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private static readonly object IntentSchema = new { type = "object", properties = new { people = new { type = "array", items = new { type = "string" } }, organizations = new { type = "array", items = new { type = "string" } }, requiredKeywords = new { type = "array", items = new { type = "string" } }, optionalKeywords = new { type = "array", items = new { type = "string" } }, receivedAfterUtc = new { type = new[] { "string", "null" } }, receivedBeforeUtc = new { type = new[] { "string", "null" } }, sortPreference = new { type = new[] { "string", "null" } }, ambiguityNote = new { type = new[] { "string", "null" } } }, required = new[] { "people", "organizations", "requiredKeywords", "optionalKeywords", "receivedAfterUtc", "receivedBeforeUtc", "sortPreference", "ambiguityNote" }, additionalProperties = false };
    private static readonly object RankingSchema = new { type = "object", properties = new { candidates = new { type = "array", items = new { type = "object", properties = new { messageId = new { type = "string" }, score = new { type = "number" }, explanation = new { type = new[] { "string", "null" } } }, required = new[] { "messageId", "score", "explanation" }, additionalProperties = false } } }, required = new[] { "candidates" }, additionalProperties = false };
}

public sealed record RankingResponse(IReadOnlyList<RankedCandidate> Candidates)
{ public RankingResponse() : this(Array.Empty<RankedCandidate>()) { } }
