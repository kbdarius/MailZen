# MailZen Database Indexing Diagnostic Record and Recovery Plan

Last updated: 2026-08-02

## Purpose

This is the single source of truth for diagnosing why MailZen has not indexed Outlook email. It records what was observed, what was tried, what failed, what was changed, and the exact order of the next tests. Future diagnostic work should update this document after each test instead of repeating earlier experiments.

## Current verified state

- Database: `%LOCALAPPDATA%\MailZen\MailZen.db`
- Diagnostic logs: `%LOCALAPPDATA%\EmailManage\diagnostic*.log`
- Current database counts, checked on 2026-08-02:
  - `messages`: 0
  - `accounts`: 0
  - `folders`: 0
  - `sync_state`: 12
- The database does **not** currently contain one year of email. It contains no indexed messages.
- The latest headless run ended with exit code `30`, meaning Outlook was unavailable.
- Outlook Classic currently cannot open normally or in Safe Mode. It reports: `Outlook has exhausted all shared resources, please close all messaging applications and restart Outlook`.
- Because Outlook itself cannot initialize MAPI, no MailZen indexing result is valid until Windows has been restarted and Outlook can open normally.
- The installer currently in `installer` is version 2.0.1, but it does not contain every diagnostic fix made afterward. Do not use it to validate the current source.
- Source now includes schema version 2 sync coverage tracking and account-level indexed range display. The installer has not yet been rebuilt for these changes.

## Important architecture facts

- MailZen reads Outlook Classic through the Outlook COM/MAPI object model.
- The indexer is read-only. It does not move, delete, modify, categorize, or mark email as read.
- A manual date-range load and a headless `--sync` load use the same indexing coordinator.
- Scheduled runs use the oldest successful folder checkpoint so missed runs can catch up after the computer was offline.
- Messages are upserted by a stable message key, so overlapping date ranges are intended to update existing rows rather than insert duplicates.
- Only Inbox folders are currently enumerated by the adapter. Archive stores may appear as selectable stores, but the adapter still asks each accessible store for its default Inbox.

## Attempts and findings

### 1. Confirmed the search problem was an empty index

The conversational search returned zero results for a moving-quote query. Direct database inspection showed `messages = 0`. The problem was therefore upstream indexing, not merely OpenAI interpretation or full-text ranking.

Result: useful diagnosis, but no messages were loaded.

### 2. Corrected the Outlook folder lookup API

The original reader attempted to call `GetFolderFromID` on a Store object. That method belongs to the MAPI Namespace. The code now calls:

`session.GetFolderFromID(folder.EntryId, folder.AccountId)`

Result: fixed a real code defect that had prevented folder reading, but indexing still encountered later failures.

### 3. Handled non-mail Outlook items and optional properties

Some Inbox collections contain reports, meeting items, calendar-related objects, or other COM items that do not expose normal mail properties. Failures observed included:

- `'System.__ComObject' does not contain a definition for 'ReceivedTime'`
- `'System.__ComObject' does not contain a definition for 'InternetMessageID'`

The reader now checks `MessageClass`, skips unsupported items, and reads optional `InternetMessageID` and `ConversationID` properties defensively.

Result: fixed real per-item compatibility defects, but did not produce a completed load before Outlook resources became exhausted.

### 4. Applied the date filter in Outlook before enumerating items

The requested date range is now translated into an Outlook `Items.Restrict` filter before sorting and walking the collection. A second date check remains in managed code.

Result: reduces the amount of COM work for a historical load. It has not yet been validated with successfully inserted messages.

### 5. Added manual and headless historical date ranges

Headless arguments now support:

- `--sync`
- `--sync-from=YYYY-MM-DD`
- `--sync-to=YYYY-MM-DD`
- `--sync-account=ACCOUNT_OR_STORE_DISPLAY_NAME` (diagnostic single-account scope)

The UI also supports a user-selected historical range.

Result: the requested one-year window reached the coordinator, as confirmed in the log, but no message rows were committed.

### 6. Added resumable checkpoints and duplicate-safe upserts

Folder sync state records successful and failed attempts. Scheduled sync uses the earliest successful checkpoint across selected accounts, and the message table uses an upsert on `message_id`.

Result: the design supports overlap and catch-up. This behavior still needs an end-to-end test with actual messages.

### 7. Allowed Outlook to be attached or started

The adapter first tries to attach to an active Outlook instance and otherwise creates `Outlook.Application`.

Result: improved unattended behavior, but starting Outlook cannot bypass a broken or exhausted MAPI session.

### 8. Removed redundant `session.Logon` calls

Explicit logon calls were removed because Outlook already owns the active profile and the extra calls could block.

Result: eliminated one possible blocking point; it did not clear the machine's existing resource-exhaustion condition.

### 9. Fixed the WPF headless-startup deadlock

The original startup path synchronously blocked the WPF dispatcher while waiting for an asynchronous headless sync. The STA Outlook work completed, but its continuation could not resume. Startup is now asynchronous and awaits the sync normally.

Result: confirmed improvement. Logs progressed beyond account discovery into indexing after this change.

### 10. Tested aggressive COM release

`Marshal.FinalReleaseComObject` was tried to force COM cleanup.

Observed failure: `COM object that has been separated from its underlying RCW cannot be used.`

This approach was reverted to a single guarded `Marshal.ReleaseComObject` call.

Result: did not work and must not be repeated without redesigning COM ownership.

### 11. Isolated inaccessible Outlook stores

Per-store exception handling was added so one inaccessible store does not abort account discovery.

Result: MailZen discovered 9 of 11 stores and skipped two stores that returned the shared-resource error. This improved partial discovery but did not load messages.

### 12. Added detailed adapter and STA logging

Logging was added around Outlook acquisition, Namespace access, Store enumeration, cleanup, STA completion, selected date range, and headless exit code.

Result: confirmed that account discovery and STA signaling can complete, and separated earlier deadlock behavior from the later Outlook/MAPI failure.

### 13. Tried Outlook normal mode and Safe Mode

Normal Outlook remained at `Opening - Outlook`. Safe Mode also remained at `Opening - Microsoft Outlook (Safe Mode)` and showed the same shared-resource error.

Result: did not work. Safe Mode failing makes an ordinary Outlook add-in a less likely cause.

### 14. Checked for other messaging processes

Outlook, Teams, Skype/Lync, and MailZen processes were checked after applications were closed. No relevant process remained, yet Outlook still could not open.

Result: closing those processes did not clear the condition. The next recovery step is a Windows restart.

### 15. Repeated full one-year runs

Several rebuilt test executables were run against the one-year range while correcting the issues above. The latest useful run found stores and entered indexing, but subsequent attempts encountered separated-RCW errors and finally Outlook shared-resource exhaustion.

Result: repetition did not add new evidence after the Outlook error appeared and likely increased diagnostic noise. No further broad run should occur until the controlled gates below pass.

## What is fixed versus what is still unproven

Fixed in current source:

- Wrong `GetFolderFromID` owner.
- Unsafe access to properties on non-mail Outlook items.
- WPF headless-startup deadlock.
- Date-range arguments and Outlook-side date restriction.
- Duplicate-safe database upsert design.
- Resume checkpoint design.
- One inaccessible store no longer aborts discovery of every store.
- Basic Outlook/STA diagnostic logging.

Still unproven:

- At least one real Outlook message can travel from COM to the `messages` table.
- The FTS table/triggers contain and return that message.
- A second overlapping load produces no duplicate.
- All intended account/store Inboxes can be indexed.
- A full one-year run completes without exhausting Outlook resources.
- The scheduler successfully catches up from the correct checkpoint.
- The conversational moving-quote query finds the expected indexed email.

## Next diagnostic plan

The tests must be performed in this order. Do not skip directly to a one-year, all-account load.

### Gate 0: restore Outlook health

1. Restart Windows.
2. Open Outlook Classic manually.
3. Wait until a normal Inbox is visible and Outlook finishes initial send/receive activity.
4. Do not run MailZen if Outlook still shows the shared-resource error.
5. Record whether Outlook opens, the time, and any error in the test log below.

Pass condition: Outlook Classic opens normally and an email can be opened manually.

If this fails: repair the Outlook profile/data-file state before testing MailZen. Possible follow-up diagnostics are Outlook profile isolation, temporarily detaching archive PST files, checking Office bitness/build, and running Microsoft Office repair. These are not authorized automatically and should be discussed before changing the Outlook profile.

### Gate 1: establish a minimal known-message test

1. Choose one accessible account and its Inbox. The headless diagnostic command can constrain scope with `--sync-account=ACCOUNT_OR_STORE_DISPLAY_NAME`.
2. Identify one known email received within the last day; record its subject and received time without putting its body in logs.
3. Run MailZen for only that account and only a one-day range.
4. Capture the process exit code and the new diagnostic-log lines.
5. Query `accounts`, `folders`, `messages`, `sync_state`, and `messages_fts` immediately afterward.

Pass condition: at least one message exists in `messages`, the known subject is present, and the corresponding FTS row exists.

If this fails: stop. Instrument the exact folder/read/write stages with counters and HRESULT/exception details before another run. Do not run a wider date range.

### Gate 2: prove duplicate protection

1. Record the message count and distinct `message_id` count.
2. Repeat the same account and one-day range.
3. Recheck both counts and inspect duplicate groups.

Pass condition: total rows equal distinct message IDs and the second run does not create duplicate rows.

### Gate 3: expand one dimension at a time

1. Run the same account for seven days.
2. If successful, run that account for 30 days.
3. Add one additional account with a one-day range.
4. Continue adding accounts individually.
5. Record duration, messages inserted/updated/skipped, and per-store failures for every run.

Pass condition: each expansion completes without Outlook/MAPI exhaustion and database coverage matches the requested window.

### Gate 4: run the one-year historical load

1. Keep Outlook open and stable.
2. Run accounts sequentially, not all stores concurrently.
3. Process bounded date slices, preferably one month at a time, while preserving overlap-safe upserts.
4. Persist progress after every folder/date slice so an interruption resumes rather than restarts.
5. Monitor Outlook resource errors and stop cleanly on the first recurrence.

Pass condition: the database has nonzero rows, min/max received dates cover the intended year for the selected accounts, no duplicate message IDs exist, and all failed folders are explicitly listed.

### Gate 5: validate search and scheduling

1. Search the database directly for likely moving-quote terms such as `moving`, `move`, `quote`, `estimate`, and sender/company clues.
2. Run the user query through local search and conversational search.
3. Verify that local-only mode never sends email content externally.
4. Simulate a missed scheduled run, then verify the next sync begins from the stored checkpoint and catches up.

Pass condition: the known moving quote is retrievable and a catch-up sync adds only new/changed messages.

### Gate 6: release only after validation

1. Run the automated test suite and a clean release build.
2. Bump the application/installer version because current source differs from installer 2.0.1.
3. Build a fresh installer and remove obsolete installer executables.
4. Install and repeat at least the Gate 1 smoke test from the installed executable.

Pass condition: installed version displays the new version, indexes a known message, and returns it in search.

## Diagnostic improvements to make before broad testing

If Gate 1 does not pass immediately after reboot, make these changes before another retry:

- Add a unique run ID to every sync log entry.
- Log per-account and per-folder start/end events.
- Log counters only: inspected, accepted, skipped non-mail, inserted, updated, and failed. Never log email bodies.
- Log COM exception HRESULTs, not only exception messages.
- Return `PartialSuccess` when some folders fail instead of returning `Success` after swallowed per-folder errors.
- Store display names alongside store IDs in database diagnostic state so failures are readable.
- Avoid performing account discovery twice in one sync run; pass the already-discovered accounts into the coordinator.
- Add a single-account/single-store diagnostic option so tests do not touch every configured PST/OST.
- Treat a headless run with zero indexed messages as `PartialSuccess`, not a successful sync.
- Use bounded date slices for historical imports and release all item/folder collections between slices.
- Display final inserted/updated/skipped/error counts in both UI and headless output.

## Test log

Append one row after every future attempt.

| Time | Build/source | Outlook health | Scope | Exit code | Messages before/after | Result and next decision |
|---|---|---|---|---:|---:|---|
| 2026-08-02 00:34 local | Current diagnostic source | Degraded; 2 of 11 stores reported shared-resource exhaustion | 2025-08-01 through 2026-08-02, discovered stores | Not cleanly completed | 0 / 0 | Do not repeat broad run; restore Outlook and begin at Gate 1. |
| 2026-08-02 00:36 local | Current diagnostic source | Outlook unavailable | Headless startup/account discovery | 30 | 0 / 0 | Windows restart required before another valid test. |
| 2026-08-02 11:13 local | MailZen 2.0.2 diagnostic build; single-account filter | Outlook opened normally | `keivandarius@zodvest.com`, 2026-08-01 through 2026-08-02 | 0 | 0 / 0 | Exit code was misleadingly 0; `sync_state` recorded separated-RCW failure. Added full folder exception logging and zero-message partial-success detection; rerun Gate 1 after rebuild. |
| 2026-08-02 11:16 local | MailZen 2.0.2 diagnostic build after COM/message-class fixes | Outlook opened normally | `keivandarius@zodvest.com`, 2026-08-01 through 2026-08-02 | 0 | 0 / 4 | Gate 1 passed: 4 Outlook mail items read, 4 `messages` rows, 4 FTS rows, successful checkpoint, no error. |
| 2026-08-02 11:17 local | Same build | Outlook opened normally | Same account/date range repeated | 0 | 4 / 4 | Gate 2 passed: count remained 4 and distinct message IDs remained 4; overlap did not duplicate rows. |
| 2026-08-02 11:18 local | Same build | Outlook opened normally | `keivandarius@zodvest.com`, 2026-07-26 through 2026-08-02 | 0 | 4 / 55 | Gate 3 seven-day expansion passed: 55 rows and 55 distinct message IDs; coverage 2026-07-27 through 2026-08-01. |
| 2026-08-02 11:20 local | Same build | Outlook opened normally | `keivandarius@zodvest.com`, 2026-07-03 through 2026-08-02 | 0 | 55 / 118 | Gate 3 thirty-day expansion passed: 118 rows, 118 distinct IDs, 118 FTS rows, coverage 2026-07-03 through 2026-08-01, and successful checkpoint with no error. |
| 2026-08-02 11:21 local | Same build | Outlook opened normally | `dariushousebills@gmail.com`, 2026-07-26 through 2026-08-02 | 0 | 118 / 147 | Second-account seven-day expansion passed: 29 new rows, 147 total and 147 distinct IDs, 147 FTS rows, successful checkpoint with no error. |

## Stop conditions

Stop the current test and diagnose before retrying if any of these occur:

- Outlook cannot open normally.
- The shared-resource error appears.
- A run hangs without a new stage/counter log for two minutes.
- The process exits successfully while one or more folders failed.
- The minimal Gate 1 run leaves `messages = 0`.
- Database count changes but FTS count does not.
- A duplicate `message_id` is detected.

## Current next action

Restart Windows, confirm Outlook Classic can open an Inbox, and then run only Gate 1: one accessible account, one Inbox, one day, and one known email. Inspect the database immediately after that single controlled run. Do not start the one-year load until Gates 1 through 3 pass.

## Sync coverage implementation validation

The sync coverage implementation was validated on 2026-08-02:

- The existing 147-message database migrated to schema version 2 without data loss.
- Account-level coverage can be queried from the local database for UI display.
- A first 30-day request read 118 Outlook items and recorded its completed interval.
- Repeating the same 30-day request read only 11 items from the two-day safety overlap.
- The database remained at 147 total messages, 147 distinct message IDs, and 147 FTS rows.
- The requested range remains unchanged from the user’s perspective; internal optimization changes only the effective Outlook read start.
