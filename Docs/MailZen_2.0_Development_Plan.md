# MailZen 2.0 Development Plan

## Document control

- Product: MailZen
- Target version: `2.0.0`
- Plan status: Approved for backlog creation; implementation not started
- Prepared: `2026-08-01`
- Repository: `kbdarius/MailZen`
- Working title: **MailZen — Conversational Outlook Search**

## 1. Executive summary

MailZen 2.0 will replace the existing email-cleanup, labeling, and Ollama-based workflow with a focused conversational search product for Microsoft Outlook Classic.

The user will be able to:

1. Index email from all Outlook Classic accounts into a private local database.
2. Select one or more likely email accounts from a sidebar.
3. Ask for an email in natural language.
4. Let MailZen translate that request into structured filters and search terms.
5. Search the local index and rank likely matches.
6. Review candidate messages with relevant excerpts and match explanations.
7. Open the original message in Outlook, with a `.msg` fallback when necessary.

The new product is intentionally search-first. It will not label, move, delete, categorize, or create Outlook rules as part of the normal workflow.

No implementation work is authorized by this document. Development begins only after the product owner explicitly instructs the developer to start.

## 2. Product vision

### Vision statement

> Ask MailZen what you remember, narrow the search to the accounts you choose, and open the right Outlook message.

### Primary user story

> As a person with several accounts connected to Outlook Classic, I want to describe an email conversationally, select the accounts where it may exist, and receive ranked candidates that I can open directly in Outlook.

### Example requests

- “Find the Wells Fargo email saying my transfer was processed.”
- “Which email from Mitch contained the CrowdStreet download?”
- “Show me the most recent message about the school tuition invoice.”
- “Find the Stryten email where someone mentioned a battery test report.”
- “Search only ZodVest and Ariya Capital for the closing documents.”

## 3. Goals and non-goals

### Goals

- Provide fast conversational search across multiple Outlook Classic accounts.
- Keep the canonical email index and full-text search database local to the PC.
- Allow the user to explicitly control which accounts, folders, and date ranges are searched.
- Use AI only where it improves natural-language understanding and result ranking.
- Avoid sending the entire mailbox or database to an AI provider.
- Support a fast/economical model and an optional smarter model.
- Incrementally synchronize new and changed messages without creating duplicates.
- Open original Outlook messages reliably whenever possible.
- Make daily synchronization observable, retryable, and safe.
- Preserve MailZen’s existing Windows desktop identity, Outlook integration knowledge, diagnostics, and installer foundation.

### Non-goals for 2.0

- Replacing Outlook as an email client.
- Sending, replying to, forwarding, moving, deleting, or categorizing messages.
- Automatically acting on instructions found inside email content.
- Training a custom AI model on the mailbox.
- Uploading the complete mailbox to OpenAI.
- Supporting New Outlook, webmail, macOS, or mobile clients in the initial release.
- Indexing attachment contents in the initial release; attachment names and metadata are in scope.
- Semantic vector search in the initial release unless full-text search plus AI reranking proves insufficient.

## 4. User experience

### 4.1 Main layout

The primary window will have three regions:

1. **Search scope sidebar**
   - Account checkboxes
   - Select all / clear all
   - Folder filters
   - Date range
   - Read/unread filter
   - Has-attachments filter
   - Synchronization status and Sync Now action

2. **Conversation and search input**
   - Natural-language prompt input
   - Search history for the current session
   - Model selector: Fast or Smart
   - Search/cancel controls

3. **Ranked candidate results**
   - Subject
   - Sender and sender address
   - Account and folder
   - Received date/time
   - Relevant excerpt
   - AI match explanation
   - Match score or confidence band
   - Open in Outlook
   - Open exported `.msg` fallback

### 4.2 Model profiles

MailZen will expose friendly profiles instead of requiring users to remember model IDs:

| Profile | Model ID | Intended use |
|---|---|---|
| Fast / Economy | `gpt-4o-mini` | Query interpretation and routine candidate reranking |
| Smart | `gpt-5.6-luna` | Ambiguous requests, harder ranking, and searches needing more reasoning |

The user can choose a default profile in Settings and override it for an individual search. MailZen must display that Smart searches can cost more. Automatic failover from Fast to Smart must be opt-in so an outage or weak result does not silently increase cost.

Model IDs must live in configuration, not be scattered through UI or service code. The provider layer must validate structured responses and tolerate a model becoming unavailable without corrupting the local index.

### 4.3 Search sequence

1. User selects accounts and optional filters.
2. User enters a natural-language request.
3. The selected AI model converts the request into a validated `SearchIntent` object.
4. MailZen combines `SearchIntent` with user-selected filters; explicit UI choices always win.
5. SQLite full-text search retrieves a bounded candidate pool.
6. Deterministic scoring narrows the pool before any candidate content is sent to the AI.
7. The selected AI model reranks a small candidate set and returns structured explanations.
8. MailZen shows the best candidates.
9. The user opens a result in Outlook or exports/opens a `.msg` fallback.

## 5. Target architecture

### 5.1 Components

```text
MailZen WPF UI
  ├── Search workspace and account filters
  ├── Settings, privacy, model, and sync controls
  └── Results and Outlook open actions

Application services
  ├── SearchOrchestrator
  ├── EmailIndexService
  ├── LocalSearchService
  ├── AiSearchService
  ├── OutlookReadService
  ├── OutlookOpenService
  ├── SyncCoordinator
  └── TaskSchedulerService

Infrastructure
  ├── SQLite + FTS5 database
  ├── Outlook Classic COM/MAPI
  ├── OpenAI Responses API
  ├── Windows Credential Manager
  └── Existing diagnostic logger

Headless synchronization
  └── MailZen.Indexer.exe or MailZen.exe --sync
```

### 5.2 Reuse from MailZen 1.x

- WPF application shell, theme, icon, and installer foundation.
- Outlook connection and account/store enumeration patterns.
- STA-thread and COM retry handling.
- Date-restricted Outlook item retrieval patterns.
- Diagnostic logging and crash reporting.
- Window state persistence.
- Release/publish configuration.

### 5.3 Retire from MailZen 1.x

- Ollama installation, startup, update, and model-download workflows.
- `OllamaClient`, `OllamaSetupService`, and their UI.
- Learning from Deleted Items.
- AI cleanup and triage workflow.
- Keep/Review/Delete/Temp scoring and categorization.
- Review-for-deletion folders and Outlook rule generation.
- Color-coding tests and category cleanup controls.
- CSV/XLSX rescoring as a primary product workflow.
- The legacy five-step cleanup wizard.

Legacy code should be removed in controlled stages after replacement functionality is tested. Removing code must never delete user email, Outlook folders, rules, or categories automatically.

## 6. Local data design

### 6.1 Database location

Recommended default:

```text
%LOCALAPPDATA%\MailZen\MailZen.db
```

The database should not be placed in Dropbox by default because mail content may be sensitive and a live SQLite database should not be continuously synchronized by a file-sync client. Backup/export can be designed separately.

### 6.2 Core tables

#### `accounts`

- `id`
- `store_id`
- `display_name`
- `email_address`
- `provider_hint`
- `is_enabled`
- `last_seen_utc`

#### `folders`

- `id`
- `account_id`
- `entry_id`
- `folder_path`
- `folder_type`
- `is_enabled`

#### `messages`

- `id`
- `account_id`
- `folder_id`
- `store_id`
- `entry_id`
- `internet_message_id`
- `conversation_id`
- `subject`
- `sender_name`
- `sender_address`
- `to_recipients`
- `cc_recipients`
- `received_utc`
- `sent_utc`
- `is_unread`
- `importance`
- `has_attachments`
- `attachment_names`
- `body_text`
- `body_hash`
- `source_modified_utc`
- `indexed_utc`
- `last_seen_utc`
- `is_missing`

#### `sync_state`

- Account/folder checkpoint
- Last successful run
- Last attempted run
- High-water timestamps
- Error summary
- Consecutive failure count

#### `search_history`

- Search text
- Selected scope
- Model profile
- Duration and result count
- Optional local-only feedback

Search history must be optional and removable from Settings.

### 6.3 Full-text search

Use SQLite FTS5 over:

- Subject
- Sender name/address
- Recipient fields
- Body text
- Attachment names

The FTS table must be synchronized transactionally with message records. Search results must support account, folder, date, read state, and attachment filters without trusting the AI to enforce them.

### 6.4 Identity and deduplication

No single Outlook identifier is sufficient for every case:

- `EntryID` can change when a message is moved.
- `StoreID + EntryID` is useful for opening the current item.
- Internet Message-ID is useful for cross-folder identity but can be absent or duplicated in edge cases.
- A body/header fingerprint provides a final fallback.

Deduplication should use a documented hierarchy and retain aliases when an item is moved so old references can resolve to the new Outlook location.

## 7. Outlook indexing strategy

### 7.1 Initial indexing

- Discover all accessible Outlook accounts/stores.
- Let the user choose accounts, folders, and initial lookback period.
- Default to Inbox with a conservative lookback period.
- Read items in date-restricted batches on a dedicated STA worker.
- Stream records into SQLite rather than holding an entire mailbox in memory.
- Commit batches transactionally.
- Expose progress and cancellation.
- Resume safely after interruption.

### 7.2 Incremental synchronization

- Query only the overlap window since the previous successful checkpoint.
- Upsert messages using the deduplication hierarchy.
- Refresh EntryID/folder aliases when messages move.
- Mark potentially missing messages conservatively; do not immediately delete index records.
- Periodically reconcile a wider time window.
- Never change Outlook read state or message contents while indexing.

### 7.3 Headless mode

The scheduled task should invoke a dedicated headless entry point rather than opening the full UI.

Recommended command:

```text
MailZen.exe --sync --quiet
```

If separating UI and indexer materially improves reliability, use `MailZen.Indexer.exe` instead. This decision should be captured in an architecture decision record before implementation.

Because Outlook COM automation requires the interactive Windows profile, the initial Task Scheduler configuration should use **Run only when the user is logged on**. MailZen must detect a busy or unavailable Outlook instance, retry safely, log the outcome, and exit with meaningful codes.

## 8. Search and AI design

### 8.1 Local-first principle

The local database performs retrieval. AI performs interpretation and bounded reranking.

The AI must not receive:

- The whole database
- Unbounded mailbox contents
- API keys or local file paths
- Content from accounts excluded by the user
- Attachments in the initial release

### 8.2 Structured search intent

The AI response should conform to a validated schema containing fields such as:

- People or organizations
- Required and optional keywords
- Subject hints
- Date interpretation
- Account hints
- Folder hints
- Attachment hints
- Sort preference
- Desired result count
- Ambiguity notes

UI-selected accounts and filters are authoritative and cannot be expanded by model output.

### 8.3 Candidate reranking

- Retrieve a configurable local candidate limit, initially 50–100.
- Narrow deterministically before AI use.
- Send only the bounded candidate metadata and short relevant excerpts.
- Require structured ranked output referencing internal candidate IDs.
- Ignore unknown candidate IDs returned by the model.
- Show why a result matched without presenting the explanation as certainty.
- Provide a Local Search Only fallback when the API is unavailable.

### 8.4 Provider abstraction

Define an interface such as `IAiSearchProvider` so model/API logic is isolated from the search pipeline. The first provider is OpenAI. The interface should support:

- Parse intent
- Rerank candidates
- Health check
- Model capability metadata
- Usage metrics
- Cancellation and timeout

No OpenAI model ID should appear in the ViewModel or XAML.

### 8.5 API key and privacy

- Use a user-provided OpenAI API key for the personal desktop release.
- Store it in Windows Credential Manager, never source control, SQLite, logs, or plain-text settings.
- Mask the key in the UI.
- Provide Test Connection and Remove Key actions.
- Redact message content from diagnostics.
- Display exactly what information may be sent for an AI-assisted search.
- Offer Local Search Only mode.
- Record token usage and estimated request cost when available.
- Add per-search and monthly budget warnings.

If MailZen is later distributed broadly, decide whether users bring their own keys or requests go through a controlled backend. A developer-owned API key must never be embedded in the desktop executable.

## 9. Opening Outlook messages

### Primary path

Use current `StoreID + EntryID` to call Outlook’s item lookup and display the original message.

### Recovery path

If the item moved and the stored EntryID is stale:

1. Search known aliases.
2. Reconcile by Internet Message-ID within the selected account.
3. Update the local location identifiers when found.

### Fallback path

Export the item to a temporary `.msg` file and provide an Open Exported Copy action. Temporary files should be named safely, expire automatically, and live under `%LOCALAPPDATA%\MailZen\Open Messages`.

## 10. Settings

Settings should include:

- Enabled accounts and folders
- Initial lookback and retention policy
- Daily sync schedule
- Default model profile
- Per-search model override behavior
- OpenAI credential status
- Local Search Only mode
- Candidate limits and result count
- Privacy disclosure
- Search-history retention
- Database size and location
- Rebuild index
- Export diagnostics without message content

Destructive maintenance actions such as deleting or rebuilding the database require explicit confirmation.

## 11. Reliability, security, and safety

- All Outlook reads occur on controlled STA workers.
- Register and revoke COM retry filters correctly.
- Release COM objects deterministically.
- Add timeouts and cancellation boundaries.
- Stream results and avoid whole-folder materialization.
- Use database transactions and migrations with backups.
- Keep a schema version table.
- Do not log subjects, bodies, recipient lists, API keys, or full prompts by default.
- Treat email bodies as untrusted data; never execute instructions found in them.
- Never allow model output to trigger email mutation or external communication.
- Keep search scope enforcement deterministic and local.
- Make all AI failures degrade to local search instead of blocking access to indexed mail.

## 12. Testing strategy

### Unit tests

- Search intent validation
- Account/filter precedence
- Deduplication and moved-message alias handling
- SQLite migrations
- FTS query construction
- Candidate scoring
- Model response validation
- Redaction and logging rules
- Cost/budget controls

### Integration tests

- SQLite indexing and search against synthetic fixtures
- Outlook adapter against a controllable abstraction/fake
- OpenAI provider with recorded/synthetic responses
- Credential storage abstraction
- Headless sync exit codes and checkpoint recovery

### Manual tests

- Each configured Outlook account type
- Outlook open, minimized, busy, closed, and unresponsive states
- Message moved after indexing
- Duplicate message across folders
- API unavailable or invalid key
- Both model profiles
- Local Search Only mode
- Large mailbox and long body performance
- Scheduled run while user is logged on

### Performance targets

Initial targets, subject to measurement:

- Normal local search results visible within 500 ms.
- AI-assisted results normally visible within 5 seconds.
- Incremental daily sync completes without blocking Outlook interaction.
- Search remains responsive with at least 250,000 indexed messages.
- Candidate payloads remain bounded regardless of database size.

## 13. Delivery phases

### Phase 0 — Architecture and safety baseline

- Freeze legacy behavior.
- Record architecture decisions.
- Add test project and service boundaries.
- Define migration and privacy contracts.

### Phase 1 — Local data foundation

- SQLite schema and migrations.
- Message/account/folder models.
- Incremental Outlook indexer.
- Deduplication and checkpoints.
- Headless sync.

### Phase 2 — Deterministic local search

- FTS5 indexing.
- Structured filters.
- Search service and performance tests.
- Local Search Only workflow.

### Phase 3 — AI-assisted search

- OpenAI provider abstraction.
- Secure credentials.
- `gpt-4o-mini` Fast profile.
- `gpt-5.6-luna` Smart profile.
- Structured intent parsing and candidate reranking.
- Usage and privacy controls.

### Phase 4 — New desktop experience

- Search/account sidebar.
- Conversational query surface.
- Ranked results and preview.
- Open in Outlook and `.msg` fallback.
- Sync and settings experience.

### Phase 5 — Automation and migration

- Task Scheduler setup and management.
- Remove legacy Ollama/cleanup UI and code.
- Preserve user data safely.
- Update installer and documentation.

### Phase 6 — Hardening and release

- Performance and failure testing.
- Privacy/security review.
- Upgrade testing from 1.x.
- Release candidate and `2.0.0` release notes.

## 14. Release gates

MailZen 2.0 is not release-ready until:

- No normal search or sync action mutates Outlook messages.
- All selected-account boundaries are covered by tests.
- Database migrations are reversible or backed up.
- Duplicate and moved-message behavior is verified.
- API credentials never appear in logs or repository files.
- Local Search Only works with no API connection.
- Both model profiles pass structured-response tests.
- Search result opening works across supported account types.
- Scheduled synchronization has clear success/failure reporting.
- Legacy Ollama setup is removed from the shipping UI.
- Documentation and installer accurately describe privacy and requirements.

## 15. Risks and mitigations

| Risk | Mitigation |
|---|---|
| Outlook COM becomes busy or unresponsive | STA isolation, retry filter, timeouts, bounded batches, cancellation |
| EntryID changes after moves | Store aliases, Internet Message-ID reconciliation, fingerprints |
| AI returns invalid or invented results | Structured validation; accept only candidate IDs supplied by MailZen |
| Sensitive email content is over-shared | Local-first retrieval, bounded excerpts, clear account scope, Local Search Only |
| API usage becomes expensive | Fast default, explicit Smart selection, candidate caps, usage/budget display |
| Model IDs change over time | Central model registry and capability checks |
| Index becomes large | FTS5, indexes, batching, retention controls, performance gates |
| Scheduled task runs outside interactive Outlook profile | Run only when logged on and report actionable failure state |
| Legacy code complicates changes | Introduce service boundaries first, then remove legacy features incrementally |
| Existing user Outlook categories remain | Never remove automatically; provide a separately confirmed cleanup tool if retained |

## 16. GitHub backlog map

The implementation backlog was created on `2026-08-01`. The epic contains a checklist linking all child issues:

- [#1 — Epic: MailZen 2.0 — Conversational Outlook search](https://github.com/kbdarius/MailZen/issues/1)
- [#2 — Establish MailZen 2.0 service boundaries and architecture decisions](https://github.com/kbdarius/MailZen/issues/2)
- [#3 — Add SQLite schema, migrations, and FTS5 email index](https://github.com/kbdarius/MailZen/issues/3)
- [#4 — Refactor Outlook COM access into a read-only indexing adapter](https://github.com/kbdarius/MailZen/issues/4)
- [#5 — Implement incremental multi-account indexing and message deduplication](https://github.com/kbdarius/MailZen/issues/5)
- [#6 — Add headless synchronization mode, checkpoints, and sync diagnostics](https://github.com/kbdarius/MailZen/issues/6)
- [#7 — Implement deterministic local email search and filtering](https://github.com/kbdarius/MailZen/issues/7)
- [#8 — Integrate OpenAI Responses API behind a structured AI search provider](https://github.com/kbdarius/MailZen/issues/8)
- [#9 — Add Fast gpt-4o-mini and Smart gpt-5.6-luna model profiles](https://github.com/kbdarius/MailZen/issues/9)
- [#10 — Secure OpenAI credentials and add privacy, usage, and budget controls](https://github.com/kbdarius/MailZen/issues/10)
- [#11 — Build conversational search workspace and account-filter sidebar](https://github.com/kbdarius/MailZen/issues/11)
- [#12 — Build ranked candidate results, excerpts, and message preview](https://github.com/kbdarius/MailZen/issues/12)
- [#13 — Open original Outlook messages with moved-item recovery and .msg fallback](https://github.com/kbdarius/MailZen/issues/13)
- [#14 — Add Windows Task Scheduler setup and daily sync management](https://github.com/kbdarius/MailZen/issues/14)
- [#15 — Remove Ollama and legacy cleanup workflows with a safe 1.x transition](https://github.com/kbdarius/MailZen/issues/15)
- [#16 — Add automated tests, failure recovery, and large-mailbox performance coverage](https://github.com/kbdarius/MailZen/issues/16)
- [#17 — Update installer, documentation, privacy guidance, and MailZen 2.0 release readiness](https://github.com/kbdarius/MailZen/issues/17)

Dependencies should be respected so UI work does not hard-code unfinished storage or provider behavior.

## 17. Recommended implementation order

1. Architecture boundaries and tests
2. Database and migrations
3. Outlook read-only adapter
4. Indexer and checkpoints
5. Local search
6. Headless sync
7. OpenAI provider and credentials
8. Model profiles and reranking
9. New UI
10. Outlook open actions
11. Scheduler management
12. Legacy removal
13. Hardening, documentation, and release

## 18. Planning commit note

Recommended commit subject:

```text
docs: plan MailZen 2.0 conversational Outlook search
```

Recommended commit body:

```text
Define the MailZen 2.0 product and technical plan for local-first,
conversational search across Outlook Classic accounts.

- replace the Ollama cleanup and labeling workflow with search
- introduce SQLite/FTS5 indexing and incremental synchronization
- plan OpenAI Fast (gpt-4o-mini) and Smart (gpt-5.6-luna) profiles
- define secure credential, privacy, cost, and local-only controls
- design ranked results with Open in Outlook and .msg fallback
- map the work into a phased GitHub backlog

This planning commit does not implement the new functionality.
```

## 19. Start condition

Creating this plan and its GitHub issues does not authorize implementation. Work begins only when the product owner gives a new explicit instruction to start, ideally by naming the first issue or development phase.
