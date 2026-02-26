# EmailManage — Storage Structure & Schemas (Phase 0)

Date: 2026-02-23

---

## 1. Per-Account Storage Path

```
%LOCALAPPDATA%\EmailManage\
├── app_settings.json              ← global app config
├── diagnostic.log                 ← rolling app log
└── accounts\
    └── <account_key>\             ← e.g. keivan_at_zodvest_com
        ├── pattern_v1.yaml        ← weighted rules + safety gates
        ├── state.json             ← watermarks, model metadata
        ├── feedback_log.csv       ← append-only user feedback
        └── triage_audit.db        ← SQLite decision trace log
```

### Account Key Generation
```
email "keivan@zodvest.com"  →  key "keivan_at_zodvest_com"
```
Replace `@` with `_at_`, `.` with `_`, lowercase, strip invalid path chars.

---

## 2. File Schemas

### 2.1 `state.json`
```json
{
  "account_key": "keivan_at_zodvest_com",
  "email_address": "keivan@zodvest.com",
  "display_name": "Keivan - ZodVest",
  "provider_hint": "Exchange",
  "pattern_version": "v1",
  "first_analyzed_at": "2026-02-23T10:00:00Z",
  "last_analyzed_at": "2026-02-23T10:00:00Z",
  "last_triage_at": null,
  "last_feedback_at": null,
  "analysis_window_months": 12,
  "total_deleted_scanned": 8412,
  "total_inbox_read_scanned": 3200,
  "total_rules_generated": 14,
  "total_keep_intent_rules": 3,
  "triage_run_count": 0,
  "feedback_event_count": 0
}
```

### 2.2 `pattern_v1.yaml`
See `Data/delete_pattern_input_seed.yaml` for full schema. Key sections:
- `scoring` — thresholds and signal requirements
- `risk_gates` — domains forced to review-only
- `keep_intent_rules` — hard safety blocks
- `rules` — weighted sender/domain/content rules
- `action_policy` — precedence and outcome logic
- `learning_policy` — how feedback updates rules

### 2.3 `feedback_log.csv`
```
event_time_utc,account_id,provider,folder,message_id,thread_id,from_address,from_domain,subject,subject_hash,has_unsubscribe,matched_rule_ids,model_score,system_action,user_action
```
Actions: `delete`, `keep_read`, `restore_from_deleted`, `delete_and_learn`, `keep_and_learn`

### 2.4 `triage_audit.db` (SQLite)
```sql
CREATE TABLE triage_decisions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id          TEXT NOT NULL,
    run_timestamp   TEXT NOT NULL,
    message_id      TEXT NOT NULL,
    from_address    TEXT,
    from_domain     TEXT,
    subject         TEXT,
    score           REAL NOT NULL,
    matched_rules   TEXT,           -- JSON array of rule IDs
    outcome         TEXT NOT NULL,  -- 'delete_candidate', 'needs_review', 'kept'
    keep_intent_hit TEXT,           -- rule ID that blocked, or NULL
    risk_gate_hit   INTEGER DEFAULT 0,
    moved_to_folder TEXT,
    restored        INTEGER DEFAULT 0,
    restored_at     TEXT
);

CREATE INDEX idx_triage_run ON triage_decisions(run_id);
CREATE INDEX idx_triage_outcome ON triage_decisions(outcome);
```

---

## 3. Event Log Schema

### Diagnostic log events (Serilog structured)
| Event | Fields |
|---|---|
| `OutlookConnected` | `timestamp`, `account_count`, `elapsed_ms` |
| `OutlookConnectionFailed` | `timestamp`, `error_type`, `message`, `stack_trace` |
| `AccountSelected` | `timestamp`, `account_key`, `email_address` |
| `AnalysisStarted` | `timestamp`, `account_key`, `window_months` |
| `AnalysisCompleted` | `timestamp`, `account_key`, `deleted_count`, `inbox_count`, `rules_count` |
| `TriageStarted` | `timestamp`, `account_key`, `run_id` |
| `TriageCompleted` | `timestamp`, `run_id`, `delete_candidates`, `needs_review`, `kept` |
| `FeedbackRecorded` | `timestamp`, `account_key`, `action`, `message_id`, `rule_ids` |
| `RuleUpdated` | `timestamp`, `rule_id`, `field`, `old_value`, `new_value` |
| `ErrorOccurred` | `timestamp`, `context`, `error_type`, `message` |

---

## 4. App Settings (`app_settings.json`)
```json
{
  "last_selected_account_key": "keivan_at_zodvest_com",
  "theme": "system",
  "log_level": "Information",
  "auto_connect_on_startup": true,
  "triage_confirmation_required": true,
  "max_triage_batch_size": 500
}
```
