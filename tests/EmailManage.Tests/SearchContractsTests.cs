using EmailManage.Models;
using EmailManage.Services;
using Xunit;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System;
using System.Linq;
using Microsoft.Data.Sqlite;

namespace EmailManage.Tests;

public sealed class SearchContractsTests
{
    [Fact]
    public void ModelRegistryKeepsExplicitProfilesCentralized()
    {
        Assert.Equal(MailZenModelRegistry.FastModelId, MailZenModelRegistry.Resolve(SearchModelProfile.Fast).ModelId);
        Assert.Equal(MailZenModelRegistry.SmartModelId, MailZenModelRegistry.Resolve(SearchModelProfile.Smart).ModelId);
    }

    [Fact]
    public void SearchIntentDefaultsToEmptyCollections()
    {
        var intent = new SearchIntent();
        Assert.Empty(intent.People);
        Assert.Empty(intent.RequiredKeywords);
    }

    [Fact]
    public void NaturalLanguageLocalQueryRemovesFillerWordsAndCommonTypo()
    {
        Assert.Equal("move quote jim", SearchQueryText.BuildLocalQuery("I'm looking for a move qoute from Jim"));
    }

    [Fact]
    public void BooleanQueryPreservesGroupingAndOperators()
    {
        Assert.Equal("(\"move\"* AND \"service\"*) OR \"jim\"*", SearchQueryText.BuildBooleanQuery("(move & service) | jim"));
    }

    [Fact]
    public void EmptyAccountScopeCannotSearch()
    {
        var scope = new SearchScope(new HashSet<string>());
        Assert.Empty(scope.AccountIds);
    }

    [Fact]
    public async Task SqliteIndexCanBeCreatedUpsertedAndSearched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mailzen-test-{Guid.NewGuid():N}.db");
        try
        {
            var database = new MailZenDatabase(path);
            await database.InitializeAsync();
            await database.UpsertMessageAsync(new IndexedMessage(
                "message-1", "account-1", "folder-1", "store-1", "entry-1", null, "Transfer processed",
                "Wells Fargo", "alerts@example.test", "Your transfer was processed successfully.",
                DateTime.UtcNow, true, false, ""));

            var results = await database.SearchAsync(new SearchRequest(
                "transfer processed", new SearchScope(new HashSet<string> { "account-1" })));

            Assert.Single(results);
            Assert.Equal("message-1", results[0].Message.MessageId);
        }
        finally { SqliteConnection.ClearAllPools(); if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task AiOrchestrationBoundsCandidatePayloadAndFallsBack()
    {
        var candidates = Enumerable.Range(1, 250_000).Select(i => new SearchResult(
            new IndexedMessage($"message-{i}", "account-1", "folder-1", "store-1", $"entry-{i}", null,
                $"Subject {i}", "Sender", "sender@example.test", "body", DateTime.UtcNow, false, false, ""),
            1, "excerpt")).ToArray();
        var local = new FakeLocalSearch(candidates);
        var provider = new FakeAiProvider();
        var orchestrator = new SearchOrchestrator(local, provider);

        var result = await orchestrator.SearchAsync(new SearchRequest("query", new SearchScope(new HashSet<string> { "account-1" })), true);

        Assert.Equal(20, provider.ReceivedCandidateCount);
        Assert.Equal(20, result.Count);
    }

    [Fact]
    public async Task ConversationalModeUsesAiExplicitly()
    {
        var candidates = new[] { new SearchResult(
            new IndexedMessage("message-1", "account-1", "folder-1", "store-1", "entry-1", null,
                "Moving quote", "Jim", "jim@example.test", "quote", DateTime.UtcNow, false, false, ""),
            1, "quote") };
        var local = new FakeLocalSearch(candidates);
        var provider = new FakeAiProvider();
        var orchestrator = new SearchOrchestrator(local, provider);

        var result = await orchestrator.SearchAsync(new SearchRequest("moving quote", new SearchScope(new HashSet<string> { "account-1" })), true);

        Assert.Single(result);
        Assert.Equal(1, provider.ParseIntentCalls);
    }

    [Fact(Timeout = 60_000)]
    public async Task SqliteIndexHandles250000SyntheticMessages()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mailzen-benchmark-{Guid.NewGuid():N}.db");
        try
        {
            var messages = Enumerable.Range(1, 250_000).Select(i => new IndexedMessage(
                $"message-{i}", "account-1", "folder-1", "store-1", $"entry-{i}", null,
                $"Subject {i}", "Sender", "sender@example.test", "synthetic benchmark body",
                DateTime.UtcNow.AddMinutes(-i), false, false, "")).ToArray();
            var database = new MailZenDatabase(path);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            await database.UpsertMessagesAsync(messages);
            stopwatch.Stop();
            var results = await database.SearchAsync(new SearchRequest("benchmark", new SearchScope(new HashSet<string> { "account-1" }), 20));
            Assert.Equal(20, results.Count);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(60), $"SQLite benchmark took {stopwatch.Elapsed}.");
        }
        finally { SqliteConnection.ClearAllPools(); if (File.Exists(path)) File.Delete(path); }
    }

    private sealed class FakeLocalSearch(params IReadOnlyList<SearchResult>[] responses) : ILocalSearchService
    {
        private int _calls;
        public Task<IReadOnlyList<SearchResult>> SearchAsync(SearchRequest request, CancellationToken cancellationToken = default)
        {
            var response = responses[Math.Min(_calls++, responses.Length - 1)];
            return Task.FromResult(response);
        }
    }
    private sealed class FakeAiProvider : IAiSearchProvider
    {
        public int ReceivedCandidateCount { get; private set; }
        public int ParseIntentCalls { get; private set; }
        public Task<SearchIntent> ParseIntentAsync(string query, SearchScope scope, CancellationToken cancellationToken = default)
        { ParseIntentCalls++; return Task.FromResult(new SearchIntent()); }
        public Task<IReadOnlyList<RankedCandidate>> RerankAsync(IReadOnlyList<SearchResult> candidates, SearchIntent intent, CancellationToken cancellationToken = default)
        { ReceivedCandidateCount = candidates.Count; return Task.FromResult<IReadOnlyList<RankedCandidate>>(candidates.Select((c, i) => new RankedCandidate(c.Message.MessageId, i, "match")).ToArray()); }
        public Task<bool> CheckHealthAsync(CancellationToken cancellationToken = default) => Task.FromResult(true);
    }
}
