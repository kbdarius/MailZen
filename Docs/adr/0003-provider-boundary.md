# ADR 0003: Provider and search boundaries

- Status: Accepted
- Date: 2026-08-01

## Decision

`ILocalSearchService` owns deterministic retrieval and scope enforcement. `IAiSearchProvider`
may interpret a query or rerank a bounded candidate list, but it cannot expand the selected
account scope or invent result IDs. `ISearchOrchestrator` coordinates the two.

OpenAI model identifiers belong in configuration/profile objects, not ViewModels, XAML,
or Outlook/indexing services.

## Rationale

Local enforcement keeps account boundaries correct even when an AI response is invalid,
ambiguous, unavailable, or maliciously influenced by email content.
