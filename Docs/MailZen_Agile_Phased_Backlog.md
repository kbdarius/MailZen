# MailZen Agile Phased Backlog (Redesign)

**Version:** 1.1.1 — 2026-03-10

### Release Notes (v1.1.1)
- Added: "Categories" column to the dataset extraction and Excel report.
- Fixed: UI text wrapping for the categories checkbox.

### Release Notes (v1.1.0)
- Added: "Score Existing Dataset" feature — score previously exported CSV/XLSX without re-extracting from Outlook.
- Added: Sender scoring (JunkScore, SenderReputation) and Recommendation column; optional Outlook color-category application.
- Docs: Updated user-journey and backlog to reflect scoring + feedback-loop changes.

Date: 2026-02-26
Project Mode: Agile, feedback-gated delivery

## 1. Working Agreement

1. Development is split into phases.
2. Every phase ends with a demo and review checkpoint.
3. No next-phase implementation starts until you approve the current phase.
4. Core behavior ships first, optional AI comes later.
5. Safety beats automation when there is uncertainty.

## 1.1 Product Goal and Outcomes

Goal: Build a Windows desktop companion tool that programs Outlook to do the work. It learns each account's email deletion patterns using a local LLM, stages suspected junk for user review in Outlook, and creates native Outlook rules based on actual user behavior.

Target outcomes:
- Zero-friction onboarding with automatic Outlook connection.
- Linear, wizard-like UI workflow.
- AI-first scoring using Ollama (no custom C# algorithms).
- Safe automation: Stage emails in `Review for Deletion`, learn from user actions, and create native Outlook rules.
- Cross-account learning: Use Microsoft's "Focused/Other" categorization to train the LLM for non-Microsoft accounts.

## 2. Phase Gate Rule (Mandatory)

At the end of each phase, we stop and run this checkpoint:

1. Demo working build.
2. Show UI screens and behavior.
3. Collect your feedback and change requests.
4. Convert feedback to tasks.
5. Re-estimate and then start next phase.

## 3. High-Level Plan

| Phase   | Goal                                    | Main Output                                                                  | Feedback Gate                                                   |
| ------- | --------------------------------------- | ---------------------------------------------------------------------------- | --------------------------------------------------------------- |
| Phase 0 | UX Redesign & Architecture Teardown     | New linear UI flow, removal of old algorithm code                            | Approve new UI layout and workflow steps                        |
| Phase 1 | Auto-Connect & Snapshot Foundation      | App auto-connects, reads "Focused/Other", tracks email IDs                   | Approve connection behavior and snapshot database schema        |
| Phase 2 | AI Learning & Triage (The Handoff)      | Ollama scans Inbox, moves emails to `Review for Deletion`, pauses for user   | Approve AI accuracy and the "Handoff" pause behavior            |
| Phase 3 | Confirm, Learn & Automate               | App checks user actions, moves kept emails back, creates Outlook rules       | Approve rule creation, false-positive handling, and edge cases  |
| Phase 4 | Settings & Hardening                    | Gearbox settings (Ollama updates, logs), error handling, release prep        | Approve settings UI and production-readiness                    |

## 4. Detailed Tasks By Phase

## Phase 0: UX Redesign & Architecture Teardown

### Objective
Strip out the old complex sidebar UI and custom scoring algorithms. Build the new linear, wizard-like UI shell.

### Tasks
- [ ] P0-T01 Delete `PatternEngine`, `RuleScorer`, and related YAML pattern files.
- [ ] P0-T02 Redesign `MainWindow.xaml` to remove the sidebar navigation.
- [ ] P0-T03 Implement a linear workflow UI (Step 1: Connect, Step 2: Learn, Step 3: Triage, Step 4: Review, Step 5: Automate).
- [ ] P0-T04 Add a "Gearbox" (Settings) icon to the UI for advanced options.
- [ ] P0-T05 Update `MainViewModel` to support the new linear state machine.
- [ ] P0-T06 Define the SQLite schema for the "Snapshot" database (tracking EntryIDs).

### Demo Output
- A clean, linear UI that visually represents the new workflow.
- Codebase stripped of old algorithm logic.

### Exit Criteria
- You approve the new visual layout and the step-by-step flow.

## Phase 1: Auto-Connect & Snapshot Foundation

### Objective
Make the app automatically connect to Outlook on launch, read Microsoft's "Focused/Other" property, and establish the database to track email states.

### Tasks
- [ ] P1-T01 Modify startup logic to auto-connect to Outlook without requiring a button click.
- [ ] P1-T02 Update `OutlookConnectorService` to read the `PR_INBOX_CATEGORIZED` (Focused/Other) MAPI property.
- [ ] P1-T03 Implement SQLite "Snapshot" repository to store `EntryID`, `OriginalFolder`, and `Status`.
- [ ] P1-T04 Create the `Review for Deletion` folder in Outlook if it doesn't exist.
- [ ] P1-T05 Add logic to fetch recent `Deleted Items` to prepare for AI learning.

### Demo Output
- App launches, auto-connects, and successfully reads the "Focused/Other" status of inbox emails.

### Exit Criteria
- You confirm the app connects seamlessly and correctly identifies "Other" emails in your Microsoft account.

## Phase 2: AI Learning & Triage (The Handoff)

### Objective
Integrate Ollama to learn from history and triage the Inbox, then pause the app to let the user review in Outlook.

### Tasks
- [ ] P2-T01 Update `OllamaClient` prompt to accept a summary of `Deleted Items` and "Other" emails as training context.
- [ ] P2-T02 Implement the "Triage" step: AI evaluates unread Inbox emails against the learned context.
- [ ] P2-T03 Move AI-flagged emails to the `Review for Deletion` folder.
- [ ] P2-T04 Log moved emails into the Snapshot database.
- [ ] P2-T05 Implement the "Handoff" UI state: Pause the app and prompt the user to review Outlook.

### Demo Output
- App scans Inbox, moves suspected junk to the review folder, and stops, waiting for your input.

### Exit Criteria
- You approve the AI's initial accuracy in picking junk.
- You approve the clarity of the "Handoff" prompt.

## Phase 3: Confirm, Learn & Automate

### Objective
Process the user's manual review, handle kept emails, ask for clarification on edge cases, and create Outlook rules.

### Tasks
- [ ] P3-T01 Implement "Continue" button logic: Check current location of emails in the Snapshot database.
- [ ] P3-T02 Handle Kept Emails: Move emails still in `Review for Deletion` back to the Inbox.
- [ ] P3-T03 Update the AI's "Do Not Touch" context with the kept emails to prevent loops.
- [ ] P3-T04 Detect conflicting signals (e.g., keeping some Wells Fargo, deleting others) and prompt user for clarification.
- [ ] P3-T05 Implement Outlook Rule creation via COM for confirmed deleted patterns (consolidating rules to avoid limits).
- [ ] P3-T06 Set rule action to "Move to Deleted Items".

### Demo Output
- Full loop completion: App learns from your review, moves kept items back, and creates a working Outlook rule.

### Exit Criteria
- You confirm kept emails are handled correctly (no loops).
- You confirm Outlook rules are created and function as expected.

## Phase 4: Settings & Hardening

### Objective
Move complex features to the Gearbox, ensure stability, and prepare for daily use.

### Tasks
- [ ] P4-T01 Build the Settings View (accessed via Gearbox icon).
- [ ] P4-T02 Add Ollama model update/download controls to Settings.
- [ ] P4-T03 Add diagnostic log export to Settings.
- [ ] P4-T04 Implement error handling for Outlook COM disconnects or Ollama crashes.
- [ ] P4-T05 Optimize performance for large mailboxes.
- [ ] P4-T06 Final polish and installer packaging.

### Demo Output
- Production-ready build with accessible settings and robust error handling.

### Exit Criteria
- You sign off on the final build for daily use.



## 5. Sprint Cadence and Review Rhythm

- Sprint length: 1 week.
- Mid-sprint check: 15-minute progress sync.
- End-of-sprint demo: mandatory.
- Post-demo feedback window: 24 to 48 hours.
- Next sprint starts only after your approval.

## 5.1 Baseline Technical Decisions (Single-Doc Merge)

- Stack: C# (.NET 8), WPF UI, Outlook COM integration.
- Storage per account:
  - `pattern_v1.yaml`
  - `state.json`
  - `feedback_log.csv`
  - `triage_audit.db` (SQLite)
- Suggested local path: `%LOCALAPPDATA%\EmailManage\accounts\<account_key>\`
- Safety model:
  - Never auto-delete from sender-only signal.
  - Require multi-signal confidence and keep-intent blocking.
  - Risky financial domains default to review-first.
- Outlook rules strategy:
  - Hybrid approach.
  - App remains decision engine.
  - Outlook Rules receives only deterministic low-risk subset.

## 5.2 Success Metrics

- `Delete Candidates` precision target: >97%.
- Manual cleanup time reduction target: >50%.
- `Needs Review` volume trends down month-over-month.
- Restore-from-deleted rate trends down month-over-month.

## 6. Definition of Done Per Phase

1. Feature works on real mailbox data.
2. Errors are handled with clear messages.
3. Logs capture enough detail for debugging.
4. Manual test checklist passes.
5. You approve UI and behavior.

## 7. Backlog Change Policy

1. New requests are added to backlog with priority.
2. In-progress phase scope changes only for blockers or critical UX issues.
3. Non-critical requests are queued for next phase.
4. Every scope change updates acceptance criteria.

## 8. Immediate Next Step

Start Phase 0 and deliver:

1. UI mock flow.
2. Final naming confirmation choices.
3. Architecture and data contract draft.
