# EmailManage — Safety Rules (Phase 0)

Date: 2026-02-23

---

## Non-Negotiable Safety Rules

### Rule S1: No Sender-Only Auto-Delete
**Never** auto-delete an email based solely on sender address or domain. Auto-delete requires at least 2 independent positive signals **and** at least 1 must be a content-based signal (subject/body pattern).

### Rule S2: Keep-Intent Blocks Auto-Delete
If any keep-intent rule matches (security alerts, fraud notices, payment/transfer, verification codes, tax/statement documents), the email is **always** routed to KEEP or REVIEW regardless of delete score.

### Rule S3: Risky Domains → Review Only
Financial, brokerage, and banking domains are hard-coded to review-only. Even with a high score, these never reach `Delete Candidates` unless the user has explicitly overridden the domain in their protected-sender list with an allow.

### Rule S4: Confirmation Before Bulk Delete
Any action that deletes more than 10 emails in a batch requires an explicit confirmation dialog showing count and top senders/subjects.

### Rule S5: Restore Always Available
Every triage run is logged. Users can always undo the last triage run and restore all moved messages to their original folders.

### Rule S6: Learning Only On Request
Pattern updates (weight changes, new rules) only happen when the user explicitly chooses `Delete + Learn` or `Keep + Learn`. The `Delete Only` action does NOT modify patterns.

### Rule S7: Auto-Degrade on High False Positive Rate
If the false-positive restore rate exceeds 5% over the last 50 triage decisions, the system automatically switches to review-only mode (no auto-delete candidates) until the model is manually retrained.

### Rule S8: No Silent Data Loss
The app never permanently deletes emails. It only moves them to Outlook folders (`Delete Candidates`, `Needs Review`). Permanent deletion is the user's responsibility via Outlook's native tools.

---

## Protected Lists (Phase 5)

- **Protected Senders**: emails from these addresses/domains always go to KEEP.
- **Protected Keywords**: emails containing these subject/body terms always go to REVIEW.

---

# EmailManage — Error States & User Messages (Phase 0)

---

## Error Catalog

| Code | Condition | User Message | Recovery |
|---|---|---|---|
| `E001` | Outlook not installed | "Outlook Desktop is not installed or not detected. Please install Microsoft Outlook and restart." | Show install link |
| `E002` | Outlook not running | "Outlook is not running. Please start Outlook and click Retry." | Retry button |
| `E003` | Outlook profile locked | "Outlook profile is in use by another process. Close other Outlook automation tools and click Retry." | Retry button |
| `E004` | No accounts found | "No email accounts found in Outlook. Please configure at least one account in Outlook." | Open Outlook settings link |
| `E005` | COM exception during enumeration | "Could not read Outlook accounts. Error: {detail}. Check that Outlook is responding and click Retry." | Retry + show log |
| `E006` | Account data folder not writable | "Cannot write to data folder: {path}. Check file permissions." | Show path |
| `E007` | Pattern file corrupt | "Pattern file for {account} is corrupted. You can re-run Analyze Patterns to rebuild it." | Re-analyze button |
| `E008` | Triage folder creation failed | "Could not create folder '{name}' in {account}. Outlook may be busy." | Retry |
| `E009` | Message move failed | "Failed to move {count} message(s). They remain in their original folder." | Show details + retry |
| `E010` | SQLite audit write failed | "Could not write triage audit log. Triage results are still valid but audit trail is incomplete." | Warning only |

---

## Phase Acceptance Criteria Baseline

### Phase 0 Acceptance
- [ ] UI names confirmed: `Analyze Patterns`, `Run Triage`, `Improve Model`
- [ ] Folder names confirmed: `Smart Cleanup` > `Delete Candidates`, `Needs Review`
- [ ] Architecture doc published and reviewed
- [ ] Storage schema and file formats defined
- [ ] Safety rules documented and agreed
- [ ] Error states cataloged with messages

### Phase 1 Acceptance
- [ ] App connects to Outlook Desktop via COM
- [ ] All configured accounts are listed with metadata
- [ ] User can select an active account
- [ ] Connection status indicator works (connected/disconnected/error)
- [ ] Missing Outlook handled gracefully with clear message
- [ ] Diagnostic log captures connection events
- [ ] App runs as standalone executable
