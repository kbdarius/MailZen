# ADR 0004: Read-only Outlook access

- Status: Accepted
- Date: 2026-08-01

## Decision

Indexing and search use a read-only `IOutlookReadService`. It may enumerate stores,
folders, and messages and retrieve message properties, but it may not save, move, delete,
categorize, mark read, create rules, or send messages. Opening a result is a separate
`IOutlookOpenService` boundary.

## Rationale

Separating reads from actions makes the safety guarantee testable and prevents legacy
cleanup behavior from leaking into the new search workflow.
