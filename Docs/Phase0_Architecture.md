# EmailManage — Technical Architecture (Phase 0)

Date: 2026-02-23  
Status: Phase 0 deliverable

---

## 1. System Overview

EmailManage is a Windows desktop application that connects to Outlook Desktop via COM Interop, learns per-account email deletion patterns, triages inbox mail into high-confidence delete vs. review buckets, and improves from user feedback.

```
┌─────────────────────────────────────────────────────────┐
│                    WPF Desktop App                       │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐             │
│  │ Account   │  │ Analyze  │  │ Triage    │             │
│  │ Picker    │  │ Patterns │  │ Runner    │             │
│  └─────┬────┘  └─────┬────┘  └─────┬─────┘             │
│        │              │              │                   │
│  ┌─────▼──────────────▼──────────────▼─────────────┐    │
│  │              Core Services Layer                 │    │
│  │  OutlookConnector │ PatternEngine │ TriageRouter │    │
│  │  FeedbackLogger   │ RuleScorer    │ SafetyGates  │    │
│  └─────────────────────┬───────────────────────────┘    │
│                         │                                │
│  ┌──────────────────────▼──────────────────────────┐    │
│  │              Data / Storage Layer                │    │
│  │  pattern_v1.yaml │ state.json │ feedback_log.csv│    │
│  │  triage_audit.db │ diagnostic.log               │    │
│  └─────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────┘
          │
          ▼
┌───────────────────┐
│  Outlook Desktop  │
│  (COM Interop)    │
└───────────────────┘
```

## 2. Module Breakdown

### 2.1 UI Layer (WPF / MVVM)
| Module | Responsibility |
|---|---|
| `MainWindow` | Navigation shell with sidebar and content area |
| `AccountPickerView` | Lists Outlook accounts, shows selection state |
| `AnalyzePatternsView` | Phase 2: bootstrap/update pattern files |
| `TriageView` | Phase 3: run triage, show summary |
| `ImproveModelView` | Phase 4: review bucket workflow |
| `SettingsView` | Phase 5+: protected lists, diagnostics export |

### 2.2 Services Layer
| Service | Responsibility |
|---|---|
| `OutlookConnectorService` | COM connection lifecycle, account enumeration, folder access |
| `PatternEngine` | Load/save/update `pattern_v1.yaml` rule sets |
| `RuleScorer` | Score an email against weighted rules + safety gates |
| `TriageRouter` | Create triage folders, move messages, log decisions |
| `FeedbackLogger` | Append to `feedback_log.csv`, update model on +Learn |
| `DiagnosticLogger` | Structured logging to `diagnostic.log` |

### 2.3 Data Layer
All per-account data stored under:
```
%LOCALAPPDATA%\EmailManage\accounts\<account_key>\
```

Where `<account_key>` = sanitized email address (e.g., `keivan_at_zodvest_com`).

## 3. Key Interfaces

```csharp
// Account metadata returned by OutlookConnector
public record OutlookAccountInfo(
    string DisplayName,
    string EmailAddress,
    string StoreName,
    string StoreFilePath,
    string ProviderHint,    // "Exchange", "IMAP", "POP3", "Outlook.com"
    bool IsConnected
);

// Connection result
public record ConnectionResult(
    bool Success,
    string? ErrorMessage,
    List<OutlookAccountInfo> Accounts
);
```

## 4. Data Flow: Account Connection (Phase 1)

```
User clicks "Connect" or app starts
  → OutlookConnectorService.ConnectAsync()
    → Marshal.GetActiveObject("Outlook.Application") or new Application()
    → Enumerate NameSpace.Stores
    → For each Store, extract account info
    → Return ConnectionResult with List<OutlookAccountInfo>
  → AccountPickerViewModel receives result
    → Populates account list
    → User selects active account
    → Selected account key stored in app state
```

## 5. Confirmed UI Names

| Element | Final Name |
|---|---|
| Primary action 1 | **Analyze Patterns** |
| Primary action 2 | **Run Triage** |
| Primary action 3 | **Improve Model** |
| Triage parent folder | **Smart Cleanup** |
| High-confidence subfolder | **Delete Candidates** |
| Uncertain subfolder | **Needs Review** |

## 6. Confirmed Folder Strategy

Under each Outlook account's mailbox root:
```
📁 Smart Cleanup
   📁 Delete Candidates    ← high-confidence auto-delete candidates
   📁 Needs Review          ← uncertain / risky items for manual review
```

Created automatically by `Run Triage`. Never auto-deleted; user must confirm.

## 7. Technology Stack

| Component | Technology |
|---|---|
| Runtime | .NET 8 (Windows) |
| UI Framework | WPF with MVVM pattern |
| Outlook Integration | COM Interop (`Microsoft.Office.Interop.Outlook`) |
| Pattern Storage | YAML (YamlDotNet) |
| State Storage | JSON (System.Text.Json) |
| Audit Storage | SQLite (Microsoft.Data.Sqlite) |
| Logging | Serilog → file sink |
| MVVM Support | CommunityToolkit.Mvvm |

## 8. Error Handling Strategy

See `Phase0_ErrorStates.md` for full catalog.

## 9. Safety Model

See `Phase0_SafetyRules.md` for full specification.
