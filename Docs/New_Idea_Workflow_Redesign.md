# MailZen - Simplified Workflow & Architecture Redesign

## 1. Core Philosophy
- **Serve Outlook, Don't Replace It:** MailZen is a companion tool that programs Outlook to do the work. Users should interact with their emails in Outlook, not in MailZen.
- **Zero-Friction Onboarding:** The app should automatically connect and guide the user through a linear, step-by-step workflow. No confusing menus or scattered buttons.
- **AI-First Scoring:** Eliminate the dual-scoring system (custom algorithm + LLM). Rely entirely on the local LLM (Ollama) for intelligent classification and heavy lifting.
- **Safe Automation:** MailZen never deletes emails directly. It stages them for review, learns from the user's actual deletion behavior in Outlook, and only then creates native Outlook rules for future automation.

## 2. The New User Experience (Linear Workflow)

The UI will be redesigned as a linear, wizard-like workflow with clear, sequential steps that light up as the user progresses.

### Step 1: Auto-Connect & Initialization (Hidden/Automatic)
- App launches and automatically connects to Outlook.
- Verifies Ollama/LLM status (downloads/starts if necessary).
- *UI State:* "Connecting to Outlook and initializing AI..." (Spinner)

### Step 2: Learn from History
- The AI scans the `Deleted Items` folder to understand what the user typically deletes.
- *UI State:* "Learning your deletion habits..." (Progress bar)

### Step 3: Triage Inbox (Stage for Review)
- The AI scans the `Inbox` (specifically unread/recent emails).
- Based on learned habits, it moves suspected junk/unwanted emails to a new Outlook folder: **`Review for Deletion`**.
- *UI State:* "Scanning Inbox for unwanted emails..."

### Step 4: User Action Required (The Handoff)
- MailZen pauses and prompts the user.
- *UI State:* "I found [X] emails you might want to delete. They are in the **`Review for Deletion`** folder in Outlook. Please go to Outlook, delete the ones you don't want, and leave the ones you want to keep. Come back here and click **Continue** when done."

### Step 5: Confirm, Learn & Automate
- User clicks "Continue" in MailZen.
- MailZen checks the `Deleted Items` folder to see which of the staged emails the user *actually* deleted.
- **Handling Kept Emails:** Any emails left in the `Review for Deletion` folder are automatically moved back to the `Inbox`. MailZen logs these as "False Positives" (emails the AI got wrong).
- **Breaking the Loop:** To prevent the AI from flagging these kept emails again on the next run, MailZen updates its local "Do Not Touch" list (or updates the LLM's context prompt) with the specific sender/subject patterns of the kept emails.
- **Handling Complex Cases (Clarification Prompts):** If the AI notices conflicting behavior (e.g., you delete 80% of Wells Fargo emails but keep 20%), MailZen will pause and ask you: *"I noticed you keep some Wells Fargo emails but delete others. Should I always delete them, never delete them, or keep asking you to review them?"*
- **Rule Creation:** For the confirmed deletions, MailZen creates native **Outlook Rules** to move similar emails to the `Deleted Items` folder in the future.
- *UI State:* "Learning from your choices and creating Outlook rules..." -> "Done! Outlook will now handle these automatically."

## 3. Technical Architecture Updates

### 3.1 State Tracking (The "Snapshot" System)
To avoid reprocessing the same emails, MailZen needs a robust tracking mechanism:
- **Watermarking:** Store the `EntryID` or timestamp of the last processed email in both `Inbox` and `Deleted Items`.
- **Staging Database:** When moving emails to `Review for Deletion`, log their IDs. When the user clicks "Continue", check the current location of those specific IDs to determine the user's decision (Deleted vs. Kept).

### 3.2 AI Simplification & Cross-Account Learning (The "Other" Strategy)
- Remove the custom C# pattern scoring engine (`PatternEngine`, `RuleScorer`, `pattern_v1.yaml`).
- Route all classification through the local LLM (Ollama).
- **Leveraging Microsoft's "Other" Category for Global Learning:** For Microsoft accounts, Outlook categorizes emails into "Focused" and "Other". MailZen will read this property via the Outlook COM API. 
- **Cross-Pollination:** Instead of just using the "Other" flag to sort the Microsoft account, MailZen will use the *characteristics* of those "Other" emails (e.g., specific unsubscribe link formats, promotional keywords, sender domains) to train the local LLM. This creates a "Global Junk Profile" that the LLM can then apply to your Gmail, Yahoo, and other accounts that lack Microsoft's advanced filtering. We use the Microsoft account as the "teacher" for the rest of your accounts.

### 3.3 UI Simplification
- Remove the sidebar navigation.
- Implement a single-window, step-by-step visual pipeline (e.g., a vertical timeline or horizontal progress steps).
- Move all complex settings (Ollama model selection, update checks, advanced diagnostics) behind a **Gear Icon (Settings)**.

## 4. Open Questions & Engineering Considerations

1. **Outlook Rule Limits:** Outlook has a limit on the size/number of rules (usually around 256KB total). If we create a new rule for every deleted sender/pattern, we might hit this limit.
   *Proposed Solution:* Consolidate rules. Instead of "If sender is X, delete", use "If sender is in [List of 50 senders], delete".
2. **LLM Context Window:** Feeding the entire history of deleted emails into the LLM to "learn" isn't feasible due to context limits.
   *Proposed Solution:* We still need a lightweight extraction step to summarize the *patterns* of deleted emails (e.g., "User deletes newsletters from @marketing.com, receipts from @store.com") to feed as a system prompt to the LLM when it scans the Inbox.
3. **Rule Action:** You mentioned the rule should "actually delete such emails, not pull them into any folder."
   *Clarification:* In Outlook, "Delete" usually moves it to the `Deleted Items` folder. "Permanently Delete" bypasses the trash. We should probably stick to standard "Move to Deleted Items" for safety, even for automated rules.

---
*Document Status: Draft - Pending User Review*