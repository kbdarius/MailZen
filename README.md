# MailZen 2.0

Conversational Outlook Classic search backed by a private local index.

MailZen indexes selected Outlook accounts into `%LOCALAPPDATA%\\MailZen\\MailZen.db`, then lets you describe the email you remember in natural language. Local Search Only works offline; optional OpenAI Fast (`gpt-4o-mini`) and Smart (`gpt-5.6-luna`) modes interpret and rerank a bounded candidate set.

---

## How to Run (on your own PC)

**Double-click `Build-MailZen.bat`** in the project root to build the release. It removes old release artifacts, builds and tests the app, creates the installer, leaves exactly these two versioned executables in the root, commits the result, and pushes `main` to GitHub:

- `MailZen-<version>.exe` — portable application executable.
- `MailZenSetup-<version>.exe` — installer executable.

Requires .NET 8 SDK, Inno Setup 6 (`ISCC.exe`), and Git authentication. The script removes nested executable files and compiler output after packaging.

---

## How to Distribute (to someone else)

New users don't need .NET installed — we publish a **self-contained** build that bundles everything.

### Option A: Send the installer (recommended)

1. Install [Inno Setup 6](https://jrsoftware.org/isinfo.php) on your build machine.
2. Run the build script:
   ```powershell
   .\build-installer.ps1
   ```
3. Send them `installer\MailZenSetup.exe` (~60-80 MB).  
   They double-click it, follow the wizard, and launch MailZen from the Start Menu or Desktop.

### Option B: Send a zip

1. Run the build script (Inno Setup not required):
   ```powershell
   .\build-installer.ps1 -SkipInstaller
   ```
2. Zip the `publish\` folder and send it.
3. They extract it and run `MailZen.exe`.

### What the recipient needs

- Windows 10 or 11 (64-bit)
- Microsoft Outlook Desktop (configured with at least one email account)
- Outlook Classic Desktop with the user signed in to Windows
- OpenAI API key only for optional AI-assisted search; keys are stored in Windows Credential Manager

No .NET installation required.

---

## Key Files

| File | Purpose |
|------|---------|
| `MailZen.bat` | Build-and-run script. Publishes Release single-file output and copies `MailZen.exe` to repo root before launching. |
| `build-installer.ps1` | Build & package script. Publishes the app as self-contained (no .NET needed) and optionally compiles an InnoSetup installer. Run this when you want to create a distributable. |
| `installer\MailZen.iss` | [Inno Setup](https://jrsoftware.org/isinfo.php) installer script. Defines the Windows installer wizard — app name, icon, Start Menu shortcuts, desktop shortcut, Outlook detection warning. Used automatically by `build-installer.ps1`. |
| `src\` | C# / WPF source code (.NET 8, CommunityToolkit.Mvvm). |
| `publish\` | Self-contained build output (created by `build-installer.ps1`). |
| `Docs\` | Design documents and backlog. |

---

## First-Launch Experience (for new users)

1. **Outlook connects automatically** — MailZen finds the running Outlook instance.
2. Select the Outlook accounts to index and choose Local Search Only or an AI-assisted model profile.

MailZen does not send, move, delete, categorize, or create rules for messages during search or synchronization. Legacy Outlook categories and rules from MailZen 1.x are left untouched.
3. **Learn → Triage → Review → Automate** — the wizard walks them through each step.
