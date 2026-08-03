# ADR 0002: Local database location and ownership

- Status: Accepted
- Date: 2026-08-01

## Decision

The canonical MailZen index lives at `%LOCALAPPDATA%\\MailZen\\MailZen.db`. It is not
stored in the repository or Dropbox by default. SQLite migrations are owned by the
database service and run before any indexing or search operation.

## Rationale

Mailbox content is private and a live SQLite database is not safe to synchronize through
a file-sync client. LocalAppData also matches the Windows desktop lifecycle.

## Constraints

- Email content is never written to diagnostics.
- Migration failure must leave the prior database usable or create a recoverable backup.
- CSV/XLSX exports remain derived outputs, never the source of truth.
