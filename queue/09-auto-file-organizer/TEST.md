# TEST — Auto File Organizer

## Pre-Setup
The script moves files in a Google Drive folder into subfolders by type/date/pattern.
Before testing, create a test folder in Google Drive:

1. Go to drive.google.com
2. Create a new folder called `TAK Test Folder`
3. Upload or create 3–4 mixed files inside it (e.g. a PDF, a Google Doc, an image)
4. Copy the folder URL from your browser — you'll paste it into Settings

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Folder URL or ID | [paste the TAK Test Folder URL] |
| Sorting Mode | By File Type |
| Auto Schedule | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → paste folder URL → click Set → confirm folder name appears → Save
6. TAKScripts → 🧪 Test Run (Preview) — verify it shows what would move without moving anything
7. ✅ `📊 Organizer Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ Preview rows logged showing file names and target folders, stats show PREVIEWED ≥ 1
9. TAKScripts → ▶️ Organize Now
10. ✅ Files in `TAK Test Folder` are now in typed subfolders (e.g. `📄 Documents`, `🖼 Images`)
11. ✅ Dashboard TOTAL MOVED stat updated, LAST RUN shows today's date

## PASS Criteria
- Menu appears after reload
- Folder resolves correctly from URL
- Test Run previews without moving files
- Organize Now actually moves files into subfolders
- Dashboard stats update correctly

## FAIL — note exactly what broke and at which step
