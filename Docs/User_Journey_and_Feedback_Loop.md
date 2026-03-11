# MailZen: User Journey & Feedback Loop

**Version:** 1.1.1 — 2026-03-10

_Updated to reflect the inclusion of the "Categories" column in datasets, the scoring system (JunkScore, SenderReputation), the "Score Existing Dataset" feature, and Outlook category feedback flow._

## The Core Challenge
MailZen runs outside of Outlook. When it scores an email and colors it **Red (Delete)**, but the user actually wants to keep it, they need an easy way to tell MailZen they made a mistake without leaving Outlook. 

**The Solution:** We introduce **Feedback Categories** in Outlook. If a user disagrees with a score, they simply change the email's category to `MailZen: Fix - Keep` or `MailZen: Fix - Delete`. On the next run, MailZen scans for these specific tags, learns from them, and adjusts the rules before scoring new emails.

---

## Architecture Diagram

![MailZen User Journey and Feedback Loop](user_journey_diagram.png)

---

## The Step-by-Step Experience

### Phase 1: The Initial Heavy Lift
1. **The Run:** User installs MailZen, selects the last 6 months, and hits "Start".
2. **The Output:** The engine scores 10,000 emails. It creates the 3-sheet Excel report for high-level curiosity, but more importantly, it paints the Outlook Inbox.
3. **The Inbox View:** In Outlook, the user groups their Inbox by Category. All the "Delete" emails group together in a massive red block. The "Review" ones are yellow. The safe ones are green.

### Phase 2: Giving Feedback inside Outlook
1. **The Discovery:** The user is scrolling through the red "Delete" group and spots an email from their favorite airline. It mistakenly got flagged because it had a high bulk/unsubscribe penalty.
2. **The Correction:** Right-clicking the email in Outlook, the user selects **"MailZen: Fix - Keep"** (a blue category MailZen generated). Outlook replaces the red tag with the blue feedback tag. 
3. **The Workflow:** The user doesn't need to open the MailZen app immediately. They can just tag emails as they notice mistakes throughout the week.

### Phase 3: The Weekly Incremental Run
1. **The Kickoff:** Once a week, the user opens MailZen and clicks *"Process New Emails"*.
2. **Step A: The Learning Phase:** Before looking at any new emails, MailZen does a quick search in Outlook specifically for emails tagged with `MailZen: Fix - Keep` or `MailZen: Fix - Delete`. 
3. **Step B: Adjusting the Brain:** 
   * It extracts the `SenderEmail` from those tagged emails.
   * It writes them to a local JSON file (e.g., `SenderOverrides.json`). 
   * *Example:* "marketing@favoriteairline.com" is now permanently Whitelisted (JunkScore forced to 0).
4. **Step C: Scoring New Mail:** MailZen now extracts the emails from the past 7 days. As it scores them, it checks the Override JSON first. The new system is smarter.
5. **Step D: Cleanup:** MailZen removes the "Fix" tags from the older emails (leaving them clean) and applies the standard Red/Yellow/Green categories to the newly extracted week of emails.