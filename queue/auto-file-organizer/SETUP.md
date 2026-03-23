# Auto File Organizer - Setup Guide

**Product:** Auto File Organizer by TAKScripts
**Version:** 1.0
**Store:** [takscripts.store](https://takscripts.store)

---

## What It Does

Automatically organizes files in a Google Drive folder by sorting them into labeled subfolders. Supports three sorting modes:

- **By File Type** — PDFs, Images, Documents, Spreadsheets, Videos, Audio, Archives, Other
- **By Date** — Year/Month folder structure (e.g., `2024/03 - March`)
- **By Naming Pattern** — Custom regex rules (e.g., `^INV- => Invoices`)

---

## Installation (5 minutes)

### Step 1: Create the Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it something like **"Auto File Organizer"**

### Step 2: Open the Script Editor

1. In your new sheet, go to **Extensions > Apps Script**
2. Delete any existing code in the editor
3. Paste the entire contents of `auto-file-organizer.gs`
4. Click the **Save** icon (or Ctrl+S / Cmd+S)
5. Name the project **"Auto File Organizer"**

### Step 3: Authorize the Script

1. Close the script editor tab
2. Reload your Google Sheet
3. You should see a **"🕷 TAKScripts"** menu appear (may take a few seconds)
4. Click **🕷 TAKScripts > ⚙️ Settings**
5. Google will ask you to authorize the script — click through the prompts:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** (bottom-left)
   - Click **Go to Auto File Organizer (unsafe)** — this is normal for custom scripts
   - Click **Allow**

### Step 4: Configure Your Settings

1. After authorization, the Settings sidebar will open
2. **Set your source folder:**
   - Open the Google Drive folder you want to organize
   - Copy the URL from your browser's address bar
   - Paste it into the "Folder URL or ID" field
   - Click **Set** to confirm
3. **Choose a sorting mode:**
   - **By File Type** — best for general cleanup
   - **By Date** — best for chronological archives
   - **By Naming Pattern** — best if your files follow naming conventions
4. **Optional: Enable auto-schedule** to run daily or weekly
5. Click **Save Settings**

### Step 5: Run It

- Click **🕷 TAKScripts > 🧪 Test Run (Preview)** first to see what would happen
- When satisfied, click **🕷 TAKScripts > ▶️ Organize Now** to move files

---

## Sorting Modes

### By File Type

Creates emoji-prefixed subfolders based on MIME type:

| Subfolder | File Types |
|-----------|-----------|
| 📄 Documents | Google Docs, Word, TXT, HTML, CSV, RTF |
| 📊 Spreadsheets | Google Sheets, Excel |
| 💻 Presentations | Google Slides, PowerPoint |
| 📑 PDFs | PDF files |
| 🖼 Images | JPEG, PNG, GIF, SVG, WebP, TIFF |
| 🎥 Videos | MP4, MOV, AVI, WebM, MKV |
| 🎵 Audio | MP3, WAV, OGG, FLAC, M4A |
| 🗄 Archives | ZIP, RAR, 7z, GZIP, TAR |
| 📦 Other | Anything else |

### By Date

Creates a nested year/month structure based on each file's last modified date:

```
Source Folder/
  📅 2024/
    01 - January/
    02 - February/
  📅 2025/
    03 - March/
```

### By Naming Pattern

Define custom regex rules (one per line) in the format `regex => Folder Name`:

```
^INV-       => Invoices
^RECEIPT    => Receipts
^IMG_       => Photos
^DRAFT      => Drafts
_final$     => Final Versions
```

- Rules are checked in order; first match wins
- Unmatched files fall back to file-type sorting
- Regex is case-insensitive

---

## Features

### Test Run / Preview

Click **🧪 Test Run (Preview)** to see exactly what would be moved without actually moving anything. Results are logged to the Activity Log sheet with a "Preview" status.

### Activity Log

Every file move (or preview) is recorded in the **📋 Activity Log** sheet with:

- Timestamp
- File name
- File type (MIME type)
- Source path
- Destination path
- Status (Moved / Preview / Error)

### Duplicate Handling

If a file with the same name already exists in the target subfolder, the script appends a number: `report.pdf` becomes `report (1).pdf`.

### Auto Schedule

Enable the auto-schedule toggle in Settings to run the organizer automatically:

- **Daily** — every morning at 9 AM
- **Weekly** — every Monday at 9 AM

The schedule uses Google Apps Script time-driven triggers and runs silently in the background.

---

## Permissions Required

| Permission | Why |
|-----------|-----|
| Google Drive | Read files, create folders, move files within your Drive |
| Google Sheets | Create and write to the Activity Log sheet |
| Script Triggers | Enable auto-scheduling |

The script never deletes files, never moves files outside your selected folder, and never accesses files in folders you haven't selected.

---

## Troubleshooting

**Menu doesn't appear after reload**
Wait 5-10 seconds and reload again. Google Sheets sometimes takes a moment to load custom menus.

**"Cannot access the source folder" error**
The folder may have been deleted, or sharing permissions changed. Open Settings and set the folder again.

**Files aren't being sorted**
Make sure files are in the top level of the source folder. The script does not recursively scan subfolders — it only processes files directly inside the selected folder.

**Trigger not running**
Go to **Extensions > Apps Script > Triggers** (clock icon on the left) to verify the trigger exists. If not, open Settings and re-enable the schedule toggle.

---

## Support

Visit [takscripts.store](https://takscripts.store) for more scripts, bundles, and support.

**Powered by TAKScripts · takscripts.store**
