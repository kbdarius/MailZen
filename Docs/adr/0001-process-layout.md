# ADR 0001: Process layout

- Status: Accepted
- Date: 2026-08-01

## Decision

MailZen remains a WPF desktop process for interactive search and Outlook actions. The
same executable exposes a `--sync --quiet` entry point for scheduled synchronization;
the entry point must not construct the WPF window. Outlook COM work runs on a dedicated
STA worker owned by the Outlook adapter.

## Rationale

Outlook Classic COM requires the user's interactive profile and STA access. Keeping the
headless entry point in the product executable avoids a second deployment artifact while
allowing Task Scheduler to run without showing the search window.

## Constraints

- Only the Outlook adapter may call Outlook COM.
- UI code communicates through application services and never owns COM objects.
- Sync and interactive search share the local database but use a single-process lock.
