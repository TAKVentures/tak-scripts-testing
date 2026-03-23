# Data Cleanup Wizard — Setup Guide

**By TAKScripts (TAK Ventures)**

One-click data cleanup toolkit for any Google Sheet. Remove duplicates, trim whitespace, fix formatting, standardize dates and phone numbers — all from a custom menu.

---

## Quick Setup (2 minutes)

1. Open any Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete any code in the editor
4. Copy and paste the entire contents of `data-cleanup-wizard.gs`
5. Click **Save** (disk icon or Ctrl+S)
6. Close the Apps Script editor
7. **Reload the spreadsheet** — the "🕷 TAKScripts" menu will appear in the menu bar

> First run: Google will ask you to authorize the script. Click "Advanced" > "Go to Data Cleanup Wizard" > "Allow". This is standard for all Google Apps Scripts.

---

## How to Use

### Option A: Menu Bar

Use **🕷 TAKScripts** in the menu bar for quick access to individual operations:

| Menu Item | What It Does |
|---|---|
| 🧹 Run All Cleanups | Runs every cleanup operation at once |
| 🧪 Preview Changes | Highlights cells that would change (yellow) without modifying data |
| 🔁 Remove Duplicates | Removes duplicate rows |
| ✂️ Trim Whitespace | Strips leading/trailing/extra spaces |
| 🔤 Fix Capitalization | Applies Title Case, UPPER, lower, or Sentence case |
| 📅 Standardize Dates | Converts mixed date formats to one consistent format |
| 🗑 Remove Empty Rows & Columns | Deletes blank rows and columns |
| 📞 Standardize Phone Numbers | Formats phone numbers consistently |
| 📧 Fix Email Formatting | Lowercases and trims email addresses |
| 🚫 Remove Special Characters | Strips unwanted characters |
| ⚙️ Cleanup Settings | Opens the sidebar with all options |
| ↩️ Undo (Restore Backup) | Restores data from the most recent backup sheet |

### Option B: Sidebar

Go to **🕷 TAKScripts > ⚙️ Cleanup Settings** to open the sidebar. From there you can:

- Check/uncheck individual cleanup operations
- Set the scope (entire sheet, selected range, or specific columns)
- Configure date format, phone format, capitalization style
- Specify which columns to check for duplicates
- Choose which special characters to keep
- Run cleanup or preview from the sidebar

---

## Features

### Backup & Undo

A backup sheet (named "💾 Backup [timestamp]") is automatically created before every cleanup operation. Use **↩️ Undo (Restore Backup)** to roll back.

### Preview Mode

Preview highlights cells in yellow that would change — without modifying any data. Run a cleanup to apply changes, or close and ignore.

### Scope Control

- **Entire Sheet** — processes all data on the active sheet
- **Selected Range** — processes only the cells you have selected
- **Specific Columns** — processes only the columns you specify (e.g., A, B, D)

### Duplicate Detection

By default, duplicates are detected by comparing all columns. You can specify which columns to compare in the settings (e.g., just column A for email, or columns A and B for first + last name).

### Date Parsing

Recognizes and converts these date formats:
- `MM/DD/YYYY`, `M/D/YYYY`
- `YYYY-MM-DD`
- `DD/MM/YYYY`
- `Month DD, YYYY` (e.g., March 15, 2024)
- `DD Month YYYY` (e.g., 15 March 2024)
- `M/D/YY` (two-digit year)

### Phone Number Formats

Choose your preferred format:
- `(XXX) XXX-XXXX` (default)
- `XXX-XXX-XXXX`
- `XXX.XXX.XXXX`
- `+1 (XXX) XXX-XXXX`

---

## Customization

All settings persist between sessions via Google Script Properties. Open the sidebar to adjust:

| Setting | Default | Options |
|---|---|---|
| Scope | Entire Sheet | Entire Sheet, Selected Range, Specific Columns |
| Date Format | MM/DD/YYYY | MM/DD/YYYY, YYYY-MM-DD, DD/MM/YYYY, Month DD, YYYY |
| Phone Format | (XXX) XXX-XXXX | (XXX) XXX-XXXX, XXX-XXX-XXXX, XXX.XXX.XXXX, +1 (XXX) XXX-XXXX |
| Capitalization | Title Case | Title Case, UPPERCASE, lowercase, Sentence case |
| Duplicate Columns | All | Any comma-separated column letters |
| Keep Special Chars | .,!?@#$%&()- | Any characters you want to preserve |

---

## Troubleshooting

**Menu doesn't appear?**
Reload the spreadsheet. The menu loads on open.

**Authorization error?**
Click "Advanced" > "Go to Data Cleanup Wizard" > "Allow". The script needs access to read and modify your spreadsheet.

**Dates not converting?**
The parser handles common US and ISO formats. Very unusual date formats may not be recognized. Standardize a few by hand first, then run again.

**Want to undo?**
Use **↩️ Undo (Restore Backup)** from the menu. Each cleanup creates a backup sheet you can restore from.

---

## License

Single-user license. Do not redistribute. For more scripts visit [takscripts.store](https://takscripts.store).
