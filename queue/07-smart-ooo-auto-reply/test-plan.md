# Test Plan: Smart OOO Auto-Reply v2.0

## Setup
1. Go to `script.google.com`
2. Create a new project → name it "TAKScripts Test — Smart OOO"
3. Delete existing code
4. Paste the contents of `script.gs`
5. Save (Ctrl+S)
6. Also open a new Google Sheet (this script needs a spreadsheet context)
7. In the Sheet, go to Extensions → Apps Script
8. Paste the code there instead (this gives it Spreadsheet context for the menu)
9. Save and reload the Sheet

## Test Steps

### 1. Custom Menu Appears
- **Action:** Reload the Google Sheet
- **Expected:** A "🕷 TAKScripts" menu appears in the toolbar
- **Verify in:** Google Sheets toolbar

### 2. Settings Sidebar Opens
- **Action:** Click TAKScripts → OOO Settings
- **Expected:** A sidebar opens on the right with:
  - Dark header with spider logo and "TAKScripts" branding
  - Date picker fields for Start and End dates
  - Reply Subject text input
  - Reply Message textarea
  - Skip Senders input
  - Save and Close buttons
- **Verify in:** Google Sheets sidebar

### 3. Settings Save and Persist
- **Action:** Enter test dates, a subject, and message. Click "Save Settings". Close the sidebar. Reopen it.
- **Expected:** All values you entered are still there
- **Verify in:** Google Sheets sidebar

### 4. Test Run Works
- **Action:** Click TAKScripts → Test Run
- **Expected:** An alert popup shows:
  - "🧪 TEST RUN — No replies will be sent"
  - Number of unread threads found
  - For each thread: sender, subject, and whether it would be replied to or skipped
- **Verify in:** Alert dialog

### 5. Start OOO Works
- **Action:** Make sure dates are set (start = today, end = tomorrow). Click TAKScripts → Start OOO
- **Expected:**
  - An alert confirms OOO is active with the date range
  - If there are unread emails, replies should be sent
- **Verify in:** Alert dialog + check Gmail Sent folder

### 6. Reply Log Sheet Created
- **Action:** After Start OOO runs, check the sheets
- **Expected:** A new sheet tab called "OOO Reply Log" appears with:
  - Styled header row (dark background, gold text)
  - Frozen first row
  - Columns: Timestamp, Sender Email, Sender Name, Original Subject
  - Proper column widths
- **Verify in:** Google Sheets tabs

### 7. No Duplicate Sheets
- **Action:** Run Start OOO again
- **Expected:** Only ONE "OOO Reply Log" sheet exists (not two)
- **Verify in:** Google Sheets tabs

### 8. Stop OOO Works
- **Action:** Click TAKScripts → Stop OOO
- **Expected:** The trigger is removed (check under Extensions → Apps Script → Triggers)
- **Verify in:** Apps Script Triggers page

### 9. About Dialog
- **Action:** Click TAKScripts → About TAKScripts
- **Expected:** A dialog shows with spider logo, version 2.0, link to takscripts.store
- **Verify in:** Modal dialog

## Edge Cases to Check
- What happens if you click Start OOO with no dates set? (Should show warning)
- What happens if there are no unread emails? (Test Run should say "All clear")
- Does the script skip newsletters/automated emails correctly?

## Pass Criteria
ALL 9 test steps must pass. Any failure = NEEDS-FIX status.
