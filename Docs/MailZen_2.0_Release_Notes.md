# MailZen 2.0 release notes

MailZen 2.0 changes the primary workflow from cleanup and labeling to conversational search across Outlook Classic accounts.

- Local SQLite/FTS5 index at `%LOCALAPPDATA%\\MailZen\\MailZen.db`.
- Read-only Outlook indexing and headless `MailZen.exe --sync --quiet` synchronization.
- Local Search Only, Fast `gpt-4o-mini`, and Smart `gpt-5.6-luna` profiles.
- OpenAI keys stored in Windows Credential Manager; only bounded selected candidates may be sent for AI search.
- Results open in Outlook or export to a temporary `.msg` under `%LOCALAPPDATA%\\MailZen\\Open Messages`.
- Existing Outlook categories, rules, folders, and user mail are not removed automatically during upgrade.

Outlook Classic must be installed and available in the signed-in Windows session for indexing, scheduled sync, and original-message opening. Local Search Only remains available without an API key.
