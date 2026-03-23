# Attachment Auto-Saver — Setup Guide

**By TAKScripts (TAK Ventures) · takscripts.store**

---

## What It Does

Attachment Auto-Saver scans your Gmail for emails with attachments and automatically saves them to organized Google Drive folders. No more manually downloading and filing attachments.

**Key Features:**
- Save attachments organized by sender, Gmail label, date, or file type
- Filter by file type (PDFs only, images only, etc.)
- Set max file size limits
- Skip inline images (signatures, logos, tracking pixels)
- Duplicate filename handling — never overwrites
- Full save log with clickable Drive links
- Test Run mode to preview before saving
- Scheduled auto-save (hourly, every 6 hours, or daily)

---

## Installation (5 minutes)

### Step 1: Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet
2. Name it something like **"Attachment Auto-Saver"**

### Step 2: Open the Script Editor

1. In your new sheet, go to **Extensions > Apps Script**
2. Delete any existing code in the editor (the default `myFunction` code)
3. Copy the entire contents of `attachment-auto-saver.gs` and paste it into the editor
4. Click the **Save** icon (or Ctrl/Cmd + S)
5. Name the project **"Attachment Auto-Saver"** when prompted

### Step 3: Authorize the Script

1. Close the Apps Script editor and **reload your spreadsheet**
2. You should see a new menu: **🕷 TAKScripts**
3. Click **🕷 TAKScripts > ⚙️ Settings**
4. Google will ask you to authorize the script — click **Continue**
5. Choose your Google account
6. Click **Advanced** (bottom left), then **"Go to Attachment Auto-Saver (unsafe)"**
7. Click **Allow**

> **Note:** Google shows this warning for all custom scripts. The script only accesses Gmail (read attachments), Google Drive (save files), and Google Sheets (write the log). It does not send any data externally.

### Step 4: Configure Settings

The settings sidebar will appear. Configure:

**Source:**
- **Gmail Label** — Only scan emails with this label (e.g., "Invoices"). Leave blank for all emails with attachments.
- **Search Query** — Additional Gmail search filters (e.g., `from:accounting@company.com` or `newer_than:7d`)

**Destination:**
- **Drive Folder ID** — Paste the ID from a Google Drive folder URL. Find it in the URL: `drive.google.com/drive/folders/THIS_IS_THE_ID`. Leave blank to use your Drive root.
- **Organization Mode:**
  - **By Sender** — Creates a folder per sender (e.g., "John Smith (john@example.com)")
  - **By Gmail Label** — Creates a folder per label
  - **By Date** — Creates folders like `2026 / 03 — March`
  - **By File Type** — Creates folders like "PDFs", "Images", "Documents"

**Filters:**
- **File Types** — Check specific types to save only those (e.g., only PDFs). Leave all unchecked to save everything.
- **Max File Size** — Skip attachments larger than this (in MB). Set to 0 for no limit.
- **Skip Inline Images** — Recommended ON. Filters out email signatures, logos, and tracking pixels.

**Schedule:**
- **Off** — Manual only (use the menu to run)
- **Every Hour** — Checks for new attachments hourly
- **Every 6 Hours** — Checks four times per day
- **Daily (6 AM)** — Once daily

Click **Save Settings** when done.

---

## Usage

### Manual Run

**🕷 TAKScripts > ▶️ Save Attachments Now**

Scans Gmail immediately and saves all matching attachments to Drive. Shows a summary when complete.

### Test Run

**🕷 TAKScripts > 🧪 Test Run**

Previews exactly what would be saved without actually saving anything. Use this to verify your settings are correct before running for real.

### View Save Log

**🕷 TAKScripts > 📊 View Save Log**

Opens the log sheet showing every saved file with timestamp, email subject, sender, filename, file type, size, and a clickable Drive link.

---

## Drive Folder Structure

All attachments are saved inside a root folder called **📎 TAKScripts — Saved Attachments**.

Example structures by organization mode:

**By Sender:**
```
📎 TAKScripts — Saved Attachments/
  ├── Jane Smith (jane@example.com)/
  │   ├── invoice-march.pdf
  │   └── contract.docx
  └── Accounting (accounting@company.com)/
      └── Q1-report.xlsx
```

**By Date:**
```
📎 TAKScripts — Saved Attachments/
  ├── 2026/
  │   ├── 01 — January/
  │   ├── 02 — February/
  │   └── 03 — March/
  └── 2025/
      └── 12 — December/
```

**By File Type:**
```
📎 TAKScripts — Saved Attachments/
  ├── PDFs/
  ├── Images/
  ├── Spreadsheets/
  └── Other/
```

---

## Tips

- **First run:** Consider using a narrow search query like `newer_than:30d` for your first run to avoid processing thousands of old emails at once. You can broaden it later.
- **Duplicate handling:** If a file with the same name already exists in the target folder, the script appends `(1)`, `(2)`, etc. — it never overwrites.
- **Processed tracking:** The script remembers which emails it has already processed (up to the most recent 5,000 message IDs). It will not re-save attachments from previously processed emails.
- **Re-process emails:** If you want to re-run on previously processed emails, you can reset the tracking by running the `resetProcessedIds()` function from the Apps Script editor (Extensions > Apps Script > select `resetProcessedIds` from the function dropdown > Run).
- **Rate limits:** Google Apps Script has daily quotas. Free accounts can read ~100 emails and create ~250 files per day via Apps Script. Workspace accounts have higher limits.

---

## Troubleshooting

| Issue | Solution |
|---|---|
| Menu doesn't appear | Reload the spreadsheet. If still missing, open Apps Script and run `onOpen` manually. |
| Authorization error | Re-authorize: Extensions > Apps Script > Run any function > follow the auth prompts. |
| No attachments found | Check your search query. Try `has:attachment newer_than:7d` to verify results. |
| Folder not found | Double-check the Drive folder ID. Open the folder in Drive and copy the ID from the URL. |
| Script times out | Google limits scripts to 6 minutes per run. Use a narrower search query or enable the auto-schedule to process in batches. |
| Inline images still saving | Make sure "Skip inline images" is checked in Settings. Some inline images may not match the detection heuristics. |

---

## Permissions Required

| Permission | Why |
|---|---|
| Gmail (read) | To search for emails and read attachments |
| Google Drive (read/write) | To create folders and save attachment files |
| Google Sheets (read/write) | To write the save log |
| Script triggers | To run on a schedule (if enabled) |

The script does **not** send any data externally. All data stays within your Google account.

---

## Support

Questions or issues? Visit **[takscripts.store](https://takscripts.store)** for support.

---

*Powered by TAKScripts · takscripts.store*
