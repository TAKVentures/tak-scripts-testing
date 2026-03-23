# TEST — Data Cleanup Wizard

## Pre-Setup
The script cleans data on whatever sheet is active. Add this messy test data to Sheet1
before running (paste into cells A1 onward):

| Name | Email | Phone | Date | Notes |
|---|---|---|---|---|
| name | email | phone | date | notes |
| john smith | JOHN@EXAMPLE.COM | 5551234567 | 3/5/2026 | hello world |
| JANE DOE | jane@example.com | (555) 987-6543 | 2026-02-14 | foo bar |
| john smith | JOHN@EXAMPLE.COM | 5551234567 | 3/5/2026 | hello world |
|  |  |  |  |  |
| Bob jones | bob@Example.COM | 555.111.2222 | January 15, 2026 | test   data |

Row 4 is a duplicate of row 3. Row 5 is blank. Emails are inconsistent case. Phone formats vary.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Scope | Entire Sheet |
| Date Format | MM/DD/YYYY |
| Phone Format | (XXX) XXX-XXXX |
| Capitalization | Title Case |
| Columns for dup check | (leave blank — compare all columns) |

## Test Steps
1. Open a blank Google Sheet → add the test data above to Sheet1
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Cleanup Settings → enter values above → Save
6. TAKScripts → 🧪 Preview Changes
7. ✅ Cells that would change are highlighted yellow — emails, phones, dates, duplicate row, blank row
8. TAKScripts → 🧹 Run All Cleanups
9. ✅ `💾 Backup [timestamp]` sheet created automatically
10. ✅ Data is cleaned: Title Case names, lowercase emails, consistent phone format, standardized dates
11. ✅ Duplicate row (row 4) removed, blank row removed
12. ✅ `📊 Cleanup Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
13. ✅ Dashboard shows TOTAL RUNS ≥ 1, CELLS FIXED ≥ 1
14. TAKScripts → ↩️ Undo (Restore Backup) → confirm
15. ✅ Original messy data restored from backup sheet

## PASS Criteria
- Menu appears after reload
- Preview highlights correctly
- Backup created before cleanup runs
- All cleanup operations applied correctly
- Dashboard logs the run
- Undo restores original data

## FAIL — note exactly what broke and at which step
