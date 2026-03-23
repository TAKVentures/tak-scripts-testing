# TEST — Attachment Auto-Saver

## Pre-Setup
The script saves Gmail attachments to a Google Drive folder automatically.
Before testing:
1. Go to drive.google.com → create a folder called `TAK Test Attachments`
2. Copy the folder URL from your browser — you'll paste it into Settings
3. Confirm the test account has at least one email with an attachment in Gmail.
   If not, send an email with a PDF or image attachment to tak.testingdemo@gmail.com first.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Destination Folder | [paste TAK Test Attachments folder URL] |
| Gmail Search Query | has:attachment newer_than:30d |
| Organize by | Sender |
| Max file size (MB) | 25 |
| Skip inline images | ON |
| Auto Schedule | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → 🧪 Test Run (Preview) — verify it shows what would save without saving anything
7. ✅ `📊 Attachment Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ Preview rows logged showing file names, no crash
9. TAKScripts → ▶️ Save Attachments Now
10. ✅ Files appear in `TAK Test Attachments` folder in Drive (organized by sender subfolder)
11. ✅ Dashboard stats update — TOTAL FILES ≥ 1, LAST RUN shows today
12. TAKScripts → ▶️ Save Attachments Now a second time
13. ✅ No duplicate files saved (dedup working), no crash

## PASS Criteria
- Menu appears after reload
- Folder resolves correctly from URL
- Test Run previews without saving
- Save Now actually saves files to Drive
- No duplicates on second run

## FAIL — note exactly what broke and at which step
