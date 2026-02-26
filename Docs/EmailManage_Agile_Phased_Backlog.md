# EmailManage Agile Phased Backlog

Date: 2026-02-23
Project Mode: Agile, feedback-gated delivery

## 1. Working Agreement
1. Development is split into phases.
2. Every phase ends with a demo and review checkpoint.
3. No next-phase implementation starts until you approve the current phase.
4. Core behavior ships first, optional AI comes later.
5. Safety beats automation when there is uncertainty.

## 1.1 Product Goal and Outcomes
Goal: Build a Windows desktop tool that learns each account's email deletion patterns, triages inbox mail into high-confidence delete vs review buckets, and improves from feedback.

Target outcomes:
- Connect to Outlook Desktop and list configured accounts.
- Create/update per-account pattern files.
- Auto-create triage folders and route candidates.
- Provide `Improve Model` workflow with optional learning.
- Keep Outlook-native rule sync optional and safety-scoped.

## 2. Phase Gate Rule (Mandatory)
At the end of each phase, we stop and run this checkpoint:
1. Demo working build.
2. Show UI screens and behavior.
3. Collect your feedback and change requests.
4. Convert feedback to tasks.
5. Re-estimate and then start next phase.

## 3. High-Level Plan

| Phase | Goal | Main Output | Feedback Gate |
|---|---|---|---|
| Phase 0 | UX and technical foundation | Clickable UI flow + architecture decisions | Approve naming, layout, and flow |
| Phase 1 | Outlook account connection | App lists Outlook accounts and selects active account | Approve account list UI and connection behavior |
| Phase 2 | First-time pattern bootstrap | `Analyze Patterns` creates initial pattern file from 12-month data | Approve analysis summary and pattern quality |
| Phase 3 | Incremental updates + triage folders | `Run Triage` creates folders and routes emails | Approve folder names, triage placement, and confidence behavior |
| Phase 4 | Improve workflow | `Improve Model` with `Delete + Learn`, `Delete Only`, `Keep + Learn` | Approve review UX and learning actions |
| Phase 5 | Safety and hardening | Undo, guardrails, logs, reliability improvements | Approve production-readiness checklist |
| Phase 6 (Optional) | Local AI assist with Ollama | AI explanation/summarization for review bucket | Approve AI usefulness and model choice |

## 4. Detailed Tasks By Phase

## Phase 0: UX + Architecture Baseline
### Objective
Lock UI naming, folder naming, and system architecture before coding core workflows.

### Tasks
- [ ] P0-T01 Confirm final button names (`Analyze Patterns`, `Run Triage`, `Improve Model`).
- [ ] P0-T02 Confirm final folder names (`Smart Cleanup`, `Delete Candidates`, `Needs Review`) or your custom names.
- [ ] P0-T03 Build low-fidelity UI mock for main screens.
- [ ] P0-T04 Build clickable prototype flow for end-to-end navigation.
- [ ] P0-T05 Define per-account storage structure and filenames.
- [ ] P0-T06 Define event log schema for `delete`, `keep_read`, `restore_from_deleted`.
- [ ] P0-T07 Define error states and user-facing messages.
- [ ] P0-T08 Define non-negotiable safety rules (no sender-only auto-delete).
- [ ] P0-T09 Publish technical architecture doc (modules, interfaces, data flow).
- [ ] P0-T10 Publish phase acceptance criteria baseline.

### Demo Output
- Clickable UX prototype.
- Architecture diagram and data contract draft.

### Exit Criteria
- You approve UI structure and naming.
- You approve folder strategy and learning labels.

## Phase 1: Outlook Connectivity + Account Picker
### Objective
Connect to Outlook Desktop profile and provide account selection UI.

### Tasks
- [ ] P1-T01 Create .NET solution and project structure.
- [ ] P1-T02 Implement Outlook COM connector service.
- [ ] P1-T03 Enumerate stores/accounts from Outlook profile.
- [ ] P1-T04 Build account list UI with active account selection.
- [ ] P1-T05 Add account metadata display (address, store name, provider hint).
- [ ] P1-T06 Add connection status indicator and retry action.
- [ ] P1-T07 Handle missing Outlook / locked profile gracefully.
- [ ] P1-T08 Add diagnostic log for connection events.
- [ ] P1-T09 Write integration smoke test for account enumeration.
- [ ] P1-T10 Package runnable internal build for review.

### Demo Output
- Working app that lists all Outlook accounts and lets you select one.

### Exit Criteria
- You confirm account list correctness.
- You approve screen layout and behavior.

## Phase 2: First-Time Pattern Bootstrap
### Objective
Implement `Analyze Patterns` first-run behavior using 12 months of deleted + read-kept inbox emails.

### Tasks
- [ ] P2-T01 Add `Analyze Patterns` action in UI.
- [ ] P2-T02 Detect whether account pattern file exists.
- [ ] P2-T03 Query `Deleted Items` for last 12 months.
- [ ] P2-T04 Query Inbox `Read` emails for last 12 months.
- [ ] P2-T05 Apply keep grace period rule (read and kept in inbox after N days).
- [ ] P2-T06 Build feature extraction pipeline (sender/domain/subject/body markers).
- [ ] P2-T07 Build initial weighted rules with safety gates.
- [ ] P2-T08 Save per-account `pattern_v1.yaml`.
- [ ] P2-T09 Save `state.json` watermark and model metadata.
- [ ] P2-T10 Show analysis summary UI (counts, top patterns, risk notes).
- [ ] P2-T11 Write unit tests for bootstrap rule generation.
- [ ] P2-T12 Add dry-run mode for bootstrap validation.

### Demo Output
- First-run pattern file generated per selected account.
- Human-readable pattern summary in UI.

### Exit Criteria
- You approve quality of detected patterns.
- You approve summary presentation and wording.

## Phase 3: Incremental Learning + Folder Triage
### Objective
Update patterns since last watermark and route inbox emails into triage folders.

### Tasks
- [ ] P3-T01 Implement incremental data pull from `last_analyzed_at`.
- [ ] P3-T02 Include new `Deleted Items` signals and `Read-kept` signals.
- [ ] P3-T03 Update rule weights from incremental evidence.
- [ ] P3-T04 Add `Run Triage` action in UI.
- [ ] P3-T05 Ensure triage parent folder exists or create it.
- [ ] P3-T06 Ensure `Delete Candidates` subfolder exists or create it.
- [ ] P3-T07 Ensure `Needs Review` subfolder exists or create it.
- [ ] P3-T08 Apply scoring and risk gates to candidate inbox messages.
- [ ] P3-T09 Move high-confidence messages to `Delete Candidates`.
- [ ] P3-T10 Move uncertain/risky messages to `Needs Review`.
- [ ] P3-T11 Log decision trace (`score`, `matched_rule_ids`, `reason`).
- [ ] P3-T12 Add triage summary panel (moved counts and reason groups).
- [ ] P3-T13 Add rollback command for last triage run.

### Demo Output
- Live folder creation and triage routing from Inbox.
- Triage summary and explanation traces.

### Exit Criteria
- You approve triage behavior on real mailbox samples.
- You approve folder naming and move logic.

## Phase 4: Improve Model Workflow
### Objective
Build review workflow for manual decisions and controlled learning.

### Tasks
- [ ] P4-T01 Add `Improve Model` screen.
- [ ] P4-T02 Load `Needs Review` messages with explanation tags.
- [ ] P4-T03 Group review items by category and confidence band.
- [ ] P4-T04 Add per-item and per-group checkboxes.
- [ ] P4-T05 Implement `Delete + Learn`.
- [ ] P4-T06 Implement `Delete Only`.
- [ ] P4-T07 Implement `Keep + Learn`.
- [ ] P4-T08 Append feedback to `feedback_log.csv`.
- [ ] P4-T09 Update model only on `+ Learn` actions.
- [ ] P4-T10 Show post-action summary of pattern updates.
- [ ] P4-T11 Add confirmation prompts for bulk actions.
- [ ] P4-T12 Add undo support for last improve action batch.

### Demo Output
- Complete review-and-learn loop from `Needs Review`.

### Exit Criteria
- You confirm review UX is clear and fast.
- You confirm learning only happens when requested.

## Phase 5: Safety, Reliability, and Release Hardening
### Objective
Make behavior safe for daily use across multiple accounts.

### Tasks
- [ ] P5-T01 Add protected sender list UI.
- [ ] P5-T02 Add protected keyword list UI.
- [ ] P5-T03 Enforce hard keep-intent rules globally.
- [ ] P5-T04 Add false-positive rate monitor.
- [ ] P5-T05 Auto-degrade to `review-only` mode if risk threshold exceeded.
- [ ] P5-T06 Improve logging and diagnostics export.
- [ ] P5-T07 Optimize performance for large inboxes.
- [ ] P5-T08 Add startup health checks.
- [ ] P5-T09 Add backup/restore for pattern and state files.
- [ ] P5-T10 Add regression test suite and runbook.
- [ ] P5-T11 Add installer/update packaging draft.
- [ ] P5-T12 Prepare release candidate build.

### Demo Output
- Production-like build with safeguards and support tooling.

### Exit Criteria
- You sign off on safety checklist and restore/rollback behavior.

## Phase 6 (Optional): Local AI Assistant via Ollama
### Objective
Add optional local AI only for explanation and categorization, not primary delete decisions.

### Tasks
- [ ] P6-T01 Add model settings page (provider, model, timeout, token budget).
- [ ] P6-T02 Integrate local Ollama API client (`http://127.0.0.1:11434`).
- [ ] P6-T03 Add AI summary for `Needs Review` groups.
- [ ] P6-T04 Add AI-generated reason labels for borderline emails.
- [ ] P6-T05 Keep rule engine as final decision authority.
- [ ] P6-T06 Add fallback when Ollama unavailable.
- [ ] P6-T07 Add prompt logging toggle and privacy mode.
- [ ] P6-T08 Benchmark local model latency and quality.
- [ ] P6-T09 Tune defaults for your hardware.

### Model Plan
- Default: `gemma3:4b` for balanced quality/speed.
- Fast fallback: `llama3.2:1b`.
- Heavy optional batch mode: `gpt-oss:20b` only when needed.

### Exit Criteria
- You approve AI usefulness and response quality.
- You approve runtime speed on your machine.

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
