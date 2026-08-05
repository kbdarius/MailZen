# MailZen 2.0 user manual

## What MailZen does

MailZen searches Outlook Classic mail conversationally. It reads selected accounts into a private local SQLite/FTS5 index and never changes message contents, read state, folders, categories, rules, or recipients during indexing and search.

The index is stored at `%LOCALAPPDATA%\\MailZen\\MailZen.db`. It is intentionally outside Dropbox because a live SQLite database should not be synchronized by a file-sync client.

## Search modes

- Local Search Only: offline search with no API request.
- `gpt-4o-mini`: Fast/Economy interpretation and bounded reranking.
- `gpt-5-nano`: selectable lower-cost structured interpretation profile.
- `gpt-5.6-luna`: Smart profile for harder requests.

The model picker is centralized and allow-listed. AI-assisted searches send only the selected-account candidate set and short excerpts; they never send the full mailbox, database, attachments, API key, or local file paths.

## Accounts and synchronization

Select accounts in the left sidebar before searching. Account selection is authoritative and cannot be expanded by model output. Use Sync Now for a manual incremental sync. Scheduled sync invokes `MailZen.exe --sync --quiet` and runs only in the signed-in Windows profile because Outlook Classic COM requires that profile.

Each Outlook account now shows its indexed email range and message count beside the account name. “Not indexed” means MailZen has no local messages for that account yet. The displayed range is the range of email currently present in the local database, so it can be narrower than the mailbox’s actual history. Select any account and request any date range; overlapping requests are safe.

Enable **Include Sent Items** when Sent Items should be indexed and included in searches. The setting is saved for future manual and scheduled syncs. Turning it off stops new Sent Items indexing and hides already indexed sent messages from results; existing sent rows are retained so enabling it again does not require re-downloading them.

Search results show the owning Outlook account and a colored **Inbox** or **Sent Items** label. Use **Group by account** above the results to group or ungroup the current results instantly without running the search again.

MailZen records completed account/folder ranges in its local sync coverage table. For a repeated or overlapping request, it preserves the user’s requested dates, fills uncovered gaps, and rereads a two-day safety overlap at completed boundaries so newly arrived messages are captured. Existing messages are updated by their stable message ID instead of duplicated. A failed range is not treated as complete and will be retried later.

## Credentials and privacy

The optional OpenAI key is stored in Windows Credential Manager. Removing the key does not disable Local Search Only. Diagnostics redact email content and credentials by default.

## Opening messages

MailZen first opens the current Outlook item by StoreID and EntryID. If that fails, use the `.msg` fallback. Temporary exports live under `%LOCALAPPDATA%\\MailZen\\Open Messages` and can be removed without affecting Outlook or the index.

## Upgrade safety

MailZen 1.x categories, rules, folders, and user email are not removed automatically. MailZen 2.0 does not include the former Ollama cleanup workflow in the shipping UI. Uninstalling MailZen does not delete Outlook data.
