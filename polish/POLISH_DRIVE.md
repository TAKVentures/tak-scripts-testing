# POLISH_DRIVE — Google Drive Structure & Naming Reference

A professional design reference for Google Drive folder structure, naming conventions, file organization, permissions, and Apps Script implementation — for scripts that create, organize, or manage files and folders in Google Drive on behalf of TAKScripts clients.

---

## 1. Folder Structure & Hierarchy

### Core Principle

Drive works best when folders mirror how humans search, not how data is technically stored. Optimize for retrieval, not classification.

### Hierarchy Depth

- **Maximum 4 levels deep.** Beyond 4 levels, navigation becomes painful for clients and script logic becomes fragile.
- **Recommended default: 3 levels.**

```
Root Folder (Level 1)       ← created once, ID stored in Script Properties
└── Category or Type (Level 2)
    └── Period or Project (Level 3)
        └── Subtype (Level 4)  ← only if volume clearly warrants it
```

### Example Hierarchy

```
AcmeCorp — TAKScripts/         ← Level 1: root, one per client
├── 🧾 Invoices/
│   ├── 2026-01 — January/     ← Level 3: period subfolders
│   ├── 2026-02 — February/
│   └── 2026-03 — March/
├── 📊 Reports/
│   └── 2026-Q1/
├── 📄 Contracts/
├── 🗄️ Archive/
│   └── 2025/
└── ⚙️ Config/                 ← internal use, not shared with client
```

### Root Folder Placement

Every Drive folder structure created by a TAKScripts script should have a clearly identifiable root. Use this pattern:

```
[ClientName] — TAKScripts
```

Examples:
- `AcmeCorp — TAKScripts`
- `Riverside Dental — TAKScripts`
- `Johnson and Associates — TAKScripts`

The em dash `—` (`\u2014`) separates the client name from the brand identifier. This distinguishes TAKScripts-generated folders from folders the client already had. Never use a hyphen here — the em dash is the visual signal.

Create the root folder once at script initialization, then store its ID in Script Properties. Do not search by name on every run — searching is slow and breaks silently if the client renames the folder.

```javascript
function getRootFolder(clientName) {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty('ROOT_FOLDER_ID');
  if (!rootId) {
    const safeName = sanitizeName(clientName);
    const folderName = `${safeName} \u2014 TAKScripts`;
    const root = DriveApp.createFolder(folderName);
    rootId = root.getId();
    props.setProperty('ROOT_FOLDER_ID', rootId);
  }
  return DriveApp.getFolderById(rootId);
}
```

### Sub-Folder Organization Axes

Choose one organizational axis per folder level and never mix axes at the same level.

**By Type (best for mixed document types, predictable for clients):**
```
AcmeCorp — TAKScripts/
├── 🧾 Invoices/
├── 📄 Contracts/
└── 📊 Reports/
```

**By Date within Type (best for ongoing, time-series content):**
```
🧾 Invoices/
├── 2026-01 — January/
├── 2026-02 — February/
└── 2026-03 — March/
```

**By Sender / Source (best for intake and form submissions):**
```
📬 Form Submissions/
├── 2026-03-23_John-Smith/
└── 2026-03-24_Jane-Doe/
```

**By Client (best when one script serves multiple end-clients):**
```
Agency Dashboard — TAKScripts/
├── 📁 AcmeCorp/
├── 📁 Riverside Dental/
└── 📁 Johnson Associates/
```

---

## 2. Folder Naming Conventions

### Capitalization Rules

Use **Title Case** for all client-facing folder names. Capitalize the first letter of every major word.

- Correct: `Invoice Archive`, `Client Reports`, `Monthly Statements`
- Incorrect: `invoice archive`, `client reports`, `MONTHLY STATEMENTS`

Use **ALL CAPS** only for top-level system or status folders that must stand out as warnings:
- `ARCHIVE`
- `DO NOT EDIT`
- `SHARED WITH CLIENT`

### Separators in Folder Names

| Separator | Use Case |
|-----------|----------|
| **Space** | All human-readable folder names: `Invoice Archive`, `March Reports` |
| **Em dash `—`** | Root folder brand separator: `AcmeCorp — TAKScripts` |
| **En dash `–`** | Period range in folder names: `2026-01 — January` |
| **Never** | Slashes `/`, colons `:`, asterisks `*`, question marks `?`, pipes `\|` — these break Drive, URLs, and exports |

### Emoji Prefixes for Folder Types

Emoji prefixes make folders scannable at a glance. Apply them consistently or not at all — mixing emoji and non-emoji folders at the same hierarchy level looks unpolished.

| Emoji | Folder Type | Example |
|-------|-------------|---------|
| `📁` | Generic container or parent folder | `📁 Client Files` |
| `📊` | Reports, analytics, dashboards | `📊 Monthly Reports` |
| `📄` | Documents, contracts, agreements | `📄 Contracts` |
| `🧾` | Invoices, billing, receipts | `🧾 Invoices` |
| `📬` | Incoming / received submissions | `📬 Intake Forms` |
| `📤` | Outgoing / sent to client | `📤 Deliverables` |
| `✅` | Completed, approved, closed | `✅ Approved` |
| `🗄️` | Archive | `🗄️ Archive` |
| `⚙️` | Internal / system / do not touch | `⚙️ Config` |
| `🔒` | Confidential or restricted | `🔒 Sensitive Records` |

**When to use emoji:** Use emoji in client-shared folders where visual scannability matters. Omit emoji in deeply nested programmatic subfolders — they add noise at that level.

### Admin and System Folders

Prefix admin or internal folders with `_` (underscore) when you want them to sort to the top and be visually distinct:

```
_Admin/
_Config/
_Templates/
_Archive/
```

Use this pattern for internal-only folders that are not shared with the client.

### Date Prefix Format for Folders

Use **ISO 8601**: `YYYY-MM-DD` or `YYYY-MM`. This sorts lexicographically, making Drive's alphabetical sort behave chronologically.

```
2026-03-15_Campaign-Launch/
2026-Q1_Quarterly-Reports/        ← YYYY-Q# for quarter-level folders
2026-03 — March/                  ← YYYY-MM + human label for monthly parent folders
```

### Max Length

- **Folder names:** 50 characters max. If you need more, the name is doing too much work.
- **Full path length:** Keep total path depth under 200 characters to avoid issues when files are exported or synced via Drive for Desktop.

### Consistency Rules

- Pick one case style per folder level and never mix within that level.
- Avoid abbreviations unless universally understood (`Q1`, `PDF`, `HR`).
- Never use: spaces inside programmatic names, `&`, `#`, `%`, `@`, `!`, `(`, `)`, `/` inside a folder name.

---

## 3. File Naming Conventions

### Standard Pattern

```
[Date]_[Type]_[Client-or-Source]_[Description].[ext]
```

All segments separated by underscores. Within each segment, use dashes to separate words — no spaces.

Examples:
- `2026-03-23_Invoice_AcmeCorp_Q1-Services.pdf`
- `2026-03-23_Contract_RiversideDental_Annual-Agreement_v1.pdf`
- `2026-03-23_Report_JohnsonAssoc_Monthly-Summary.pdf`
- `2026-03-23_Submission_John-Smith_Intake-Form.pdf`

**Segment rules:**
- `[Date]` — Always ISO 8601: `YYYY-MM-DD`
- `[Type]` — Short noun: `Invoice`, `Contract`, `Report`, `Submission`, `Export`, `Receipt`
- `[Client-or-Source]` — CamelCase, no spaces: `AcmeCorp`, `RiversideDental`
- `[Description]` — Brief, hyphen-separated: `Q1-Services`, `Monthly-Summary`

### Date Format in File Names

ISO 8601 dates sort correctly in alphanumeric order. `2026-03-23` comes before `2026-04-01` when sorted by name. Month names and US-format dates (`03-23-2026`) both sort incorrectly.

```javascript
// Standard date format for file and folder names
const date = new Date();
const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
// Result: "2026-03-23"

// For folder names with human-readable month label
const monthFolder = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
// Result: "2026-03" — then append " — March" manually or via getMonth() lookup

// For timestamps when you need time precision
const timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
// Result: "2026-03-23_14-30"
```

**Key pattern notes:**
- `yyyy` = 4-digit year (lowercase y in Apps Script)
- `MM` = 2-digit month (uppercase M)
- `dd` = 2-digit day (lowercase d)
- `HH` = 24-hour hour (uppercase H)
- `mm` = minutes (lowercase m)

### Version Numbering

When a file goes through revisions, append a version suffix before the extension:

```
2026-03-23_Contract_AcmeCorp_Service-Agreement_v1.pdf
2026-03-23_Contract_AcmeCorp_Service-Agreement_v2.pdf
2026-03-23_Contract_AcmeCorp_Service-Agreement_FINAL.pdf
```

- Start at `v1`, never `v0`
- Use `FINAL` only once — it means no further revisions
- Never use `FINAL2` — if there is a `FINAL2`, rename the sequence: `v1`, `v2`, `v3`, `FINAL`
- When a file is approved and locked, the `FINAL` label is sufficient — do not add `_APPROVED_FINAL_USE-THIS-ONE`

### Specialized Naming Patterns

Invoices:
```
2026-03-23_Invoice_AcmeCorp_INV-2026-047.pdf
```

Receipts:
```
2026-03-23_Receipt_AcmeCorp_USD-1250.pdf
```

Contracts:
```
2026-03-23_Contract_AcmeCorp_Service-Agreement_v1.pdf
```

### Avoiding Special Characters

These characters must never appear in Drive file or folder names created by Apps Script:

| Character | Problem | Replacement |
|-----------|---------|-------------|
| `/` | Interpreted as path separator | Omit or use `-` |
| `:` | Illegal on Windows, breaks sync | Omit |
| `*` | Wildcard, breaks searches | Omit |
| `?` | Breaks URLs | Omit |
| `"` | Breaks JSON and CSV exports | Omit |
| `\|` | Breaks shell and exports | Omit |
| `\` | Breaks Windows paths | Omit |
| `&` | Encoding issues in URLs | Spell out `and` |
| `#` | URL fragment identifier | Omit |
| `%` | URL encoding character | Omit |

Sanitize any user-provided input before using it in a file or folder name:

```javascript
function sanitizeName(name) {
  return name
    .replace(/[\/\\:*?"'|<>#%&!@()\[\]{}]/g, '') // remove illegal characters
    .replace(/\s+/g, '-')                          // spaces to dashes
    .replace(/-{2,}/g, '-')                        // collapse multiple dashes
    .trim();
}
```

---

## 4. File Organization Patterns

### By Client (recommended default for multi-client scripts)

All client assets live under their root folder regardless of date or type. Best when one script serves many clients.

```
Agency Dashboard — TAKScripts/
├── AcmeCorp/
│   ├── 🧾 Invoices/
│   └── 📊 Reports/
└── BlueSky/
    ├── 🧾 Invoices/
    └── 📊 Reports/
```

### By Type Within Client

The most common TAKScripts pattern: root per client, type at Level 2, date period at Level 3.

```
AcmeCorp — TAKScripts/
├── 🧾 Invoices/
│   ├── 2026-01 — January/
│   └── 2026-02 — February/
└── 📊 Reports/
    └── 2026-Q1/
```

### By Date

Best for high-volume output scripts that generate files daily.

```
Output/2026/03/2026-03-15_Report.pdf
Output/2026/03/2026-03-16_Report.pdf
```

### By Type (flat)

Best for shared assets used across clients — templates, brand assets, legal docs.

```
_Templates/
_Brand-Assets/
_Legal-Boilerplate/
```

### Choosing a Pattern

| If your script... | Use pattern |
|---|---|
| Creates files for one specific client | By Type within a single client root |
| Creates files for many clients from one run | By Client, then by Type |
| Generates time-series data (daily, weekly) | By Date, then by Type |
| Manages internal templates and reference docs | By Type, flat |

---

## 5. Color-Coding Folders

### When Color Coding Helps vs Creates Confusion

**Use color coding when:**
- A small team manages all the folders and understands the color scheme
- The color-to-meaning mapping is documented somewhere accessible
- Colors track a clear lifecycle (in progress → review → complete → archive)
- There are fewer than 10 colored folders at any given time

**Avoid color coding when:**
- Multiple people create folders and apply colors inconsistently
- The color meaning exists only in one person's head
- You are applying colors to folders at all hierarchy levels (creates visual noise)
- The client has their own color conventions that conflict

**Recommendation:** Apply colors only to Level 2 folders. Leave Level 1 (root) and Level 3 (date/period) folders as default.

### Color Meaning Conventions

| Color | Conventional Meaning |
|-------|---------------------|
| Red | Urgent, overdue, attention required |
| Orange | In progress, pending action |
| Yellow | Under review, waiting on client |
| Green | Complete, approved, delivered |
| Blue (default) | General purpose, no special status |
| Teal | Templates, reference material |
| Purple | Archive, historical |
| Grey | Inactive, deprecated, do not use |

### Setting Folder Colors via Apps Script

`DriveApp` exposes a `Folder.setColor()` method using the `DriveApp.Color` enum:

```javascript
// Using DriveApp.Color enum (preferred — readable and stable)
folder.setColor(DriveApp.Color.GREEN);   // complete
folder.setColor(DriveApp.Color.RED);     // urgent
folder.setColor(DriveApp.Color.YELLOW);  // pending review
folder.setColor(DriveApp.Color.GREY);    // archive / inactive
folder.setColor(DriveApp.Color.TEAL);    // templates / reference
folder.setColor(DriveApp.Color.PURPLE);  // archive
folder.setColor(null);                   // reset to default (no color)

// Full list of DriveApp.Color values:
// CHOCOLATE, CORNFLOWER, DARK_CYAN, DARK_MAGENTA, DARK_ORANGE,
// DARK_PINK, DARK_PURPLE, DARK_TEAL, DARK_YELLOW, GREEN, GREY,
// INDIGO, LIME, PINK, PURPLE, RED, TEAL, YELLOW, BLUE, CYAN
```

### Status-Driven Color Helper

```javascript
function colorCodeFolder(folder, status) {
  const colorMap = {
    active:    DriveApp.Color.GREEN,
    pending:   DriveApp.Color.YELLOW,
    urgent:    DriveApp.Color.RED,
    archive:   DriveApp.Color.GREY,
    template:  DriveApp.Color.TEAL,
    inprogress: DriveApp.Color.CORNFLOWER,
  };
  folder.setColor(colorMap[status] ?? null);
}

// Usage
colorCodeFolder(archiveFolder, 'archive');
colorCodeFolder(activeProjectFolder, 'active');
```

---

## 6. Sharing Permissions

### Permission Levels

| Role | Can View | Can Comment | Can Edit | Can Share | Can Delete |
|------|----------|-------------|----------|-----------|------------|
| Viewer | Yes | No | No | No | No |
| Commenter | Yes | Yes | No | No | No |
| Editor | Yes | Yes | Yes | Yes* | Yes* |
| Owner | Yes | Yes | Yes | Yes | Yes |

*Editors can share and delete unless the owner has restricted it via `setShareableByEditors(false)`.

### Sharing Best Practices

- **Prefer folder-level sharing.** Adding individual users to individual files creates unmanageable permission sprawl.
- Apply sharing to the highest appropriate folder and let children inherit.
- Only use file-level sharing when a single file in a shared folder must be more restricted than its parent.
- Scripts should never automatically grant `ANYONE` (public internet) access. Default to `PRIVATE` or `ANYONE_WITH_LINK` at minimum.
- Separate sharing logic from creation logic — do not auto-share during batch runs without an explicit flag.

### Sharing in Apps Script

```javascript
// Share a folder with a specific user as editor
folder.addEditor('colleague@example.com');

// Share with viewer only
folder.addViewer('client@example.com');

// Share with commenter
folder.addCommenter('reviewer@example.com');

// Set link sharing — anyone with link can view (safe default for clients)
folder.setSharing(
  DriveApp.Access.ANYONE_WITH_LINK,
  DriveApp.Permission.VIEW
);

// Restrict to domain only (Workspace accounts)
folder.setSharing(
  DriveApp.Access.DOMAIN,
  DriveApp.Permission.VIEW
);

// Fully private (only explicitly added people)
folder.setSharing(
  DriveApp.Access.PRIVATE,
  DriveApp.Permission.NONE
);

// Remove a user's access
folder.removeEditor('old-colleague@example.com');
folder.removeViewer('old-client@example.com');

// Prevent editors from re-sharing (owner-only sharing)
folder.setShareableByEditors(false);
```

### Client-Facing Shared Folder Design

When a client opens a Drive folder shared with them, they see everything — every folder, every file name, every subfolder. There is no way to hide content without removing sharing permissions.

This means:
- No internal notes or draft names in shared folders
- No folders named `old`, `backup`, `DO NOT USE`, or `testing`
- No files named `COPY OF invoice.pdf` or `invoice REVISED AGAIN.pdf`
- No internal tracking spreadsheets or workflow status folders

A client-shared folder should feel like a well-organized filing cabinet designed for them:

```
📁 AcmeCorp — Your Documents
  ├── 📄 Contracts
  │     └── 2026-03-23_Contract_Service-Agreement_v1.pdf
  ├── 🧾 Invoices
  │     ├── 2026-01-01_Invoice_INV-2026-001.pdf
  │     └── 2026-02-01_Invoice_INV-2026-002.pdf
  └── 📊 Reports
        └── 2026-03-23_Report_Monthly-Summary.pdf
```

---

## 7. Template Files

### Where to Store Templates

Keep all templates in a dedicated `_Templates/` folder at the root level of your script's Drive area — never inside client or project folders.

```
_Templates/
├── Invoice-Template_v2.sheet
├── Contract-Template_v1.doc
└── Report-Template_v3.sheet
```

Store template file IDs in Script Properties — never search by name in production runs. Name lookups are slow and fail silently when someone renames the file.

```javascript
function getTemplateId(templateKey) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(`TEMPLATE_${templateKey}`);
  if (!id) throw new Error(`Template ID not found for key: ${templateKey}`);
  return id;
}
```

### Copying Templates

Never write to a template file directly. Always copy first, then write to the copy.

```javascript
function copyTemplateForClient(templateKey, destFolderId, clientName, date) {
  const templateId = getTemplateId(templateKey);
  const template = DriveApp.getFileById(templateId);
  const dest = DriveApp.getFolderById(destFolderId);
  const safeName = sanitizeName(clientName);
  const fileName = `${date}_${safeName}_Report_v1`;
  return template.makeCopy(fileName, dest);
}

// Usage
const copy = copyTemplateForClient(
  'REPORT',
  clientFolderId,
  'Acme Corp',
  '2026-03-23'
);
```

### Naming Template Copies

- Remove the word `Template` from the copy name.
- Insert the date and client name at the front.
- Always start copies at `v1`.
- The template file itself retains its `Template` label and its own version: `Report-Template_v3.sheet`

| Template file | Copy name |
|---|---|
| `Report-Template_v3.sheet` | `2026-03-23_AcmeCorp_Report_v1.sheet` |
| `Invoice-Template_v2.sheet` | `2026-03-23_AcmeCorp_Invoice-March_v1.sheet` |
| `Contract-Template_v1.doc` | `2026-03-23_AcmeCorp_Service-Agreement_v1.doc` |

---

## 8. Archive Patterns

### When to Archive

Archive on a schedule or a trigger — not ad hoc:

- **Time-based:** Move files to archive at end of each calendar year (January 1 run)
- **Status-based:** Move to archive when a contract is marked Closed or an invoice is marked Paid and older than 90 days
- **Count-based:** When a folder exceeds 200 files, archive everything older than 12 months

Never delete — always archive. The client may need old files for audits, disputes, or reference.

### Archive Folder Structure

Active and archived files must never live in the same folder. A client browsing `Invoices` should see only current-year or open invoices, not three years of history.

```
🗄️ Archive/
├── 2024/
│   ├── 🧾 Invoices/
│   └── 📊 Reports/
└── 2025/
    ├── 🧾 Invoices/
    └── 📊 Reports/
```

The Archive folder sits at Level 2 under the root. Inside it, folders are organized by year first, then by type — the reverse of the active structure. This is intentional: active content is accessed by type then drilled by date; archived content is accessed by year first because that is how people remember old records.

### Archive Naming

Prefix archived folders with the archive date so you know when, not just that, it was archived:

```
2026-03-15_ARCHIVED_Acme-Corp_Q1-Campaign/
```

```javascript
function markAsArchived(folder) {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const current = folder.getName();
  if (!current.includes('ARCHIVED')) {
    folder.setName(`${date}_ARCHIVED_${current}`);
  }
}
```

### Moving to Archive via Apps Script

```javascript
function archiveFolder(sourceFolderId, archiveRootId) {
  const source = DriveApp.getFolderById(sourceFolderId);
  const archiveRoot = DriveApp.getFolderById(archiveRootId);

  // Get or create year subfolder
  const year = new Date().getFullYear().toString();
  const existing = archiveRoot.getFoldersByName(year);
  const yearFolder = existing.hasNext() ? existing.next() : archiveRoot.createFolder(year);

  // Move the folder
  source.moveTo(yearFolder);
  markAsArchived(source);
  source.setColor(DriveApp.Color.GREY);
}
```

---

## 9. File Types — When to Use Which Format

### Sheets vs CSV

| Use Sheets when... | Use CSV when... |
|---|---|
| Data needs formulas or formatting | Data will be consumed by another system |
| Sharing with non-technical stakeholders | Size > 5MB and formatting is irrelevant |
| Multiple tabs of related data | Importing into a database |
| Ongoing editing is expected | One-time export or snapshot |
| You need version history | Recipient uses non-Google tooling |

### Docs vs PDF

| Use Docs when... | Use PDF when... |
|---|---|
| Document will be revised collaboratively | Document is final and must not be edited |
| Sharing internally with editors | Sending to clients or external parties |
| Template that will be copied and filled | Legal, compliance, or audit record |

### Native vs Export Formats

- **Keep native Google formats** (Sheets, Docs, Slides) for everything you edit.
- **Export to PDF/CSV/XLSX** only for delivery. Store exports alongside the native file, not instead of it.
- Never import a CSV, edit it as a CSV, then re-export — convert it to a Sheet at intake and work natively.
- Native Google files (Sheets, Docs, Slides, Forms) **do not count against Drive storage quota**. Uploaded files (PDF, XLSX, images) do.

### Creating Files via Apps Script

```javascript
// Create a native Google Sheet inside a folder
function createSheet(folder, name) {
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  return ss;
}

// Create a plain text file
function createTextFile(folder, name, content) {
  return folder.createFile(name, content, MimeType.PLAIN_TEXT);
}

// Export a Sheet as PDF and save it to a folder
function saveSheetAsPdf(spreadsheetId, folder, fileName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const blob = response.getBlob().setName(`${fileName}.pdf`);
  return folder.createFile(blob);
}
```

---

## 10. Search and Retrieval

### Naming for Searchability

Drive's full-text search indexes file names and document content. Optimize names for how a user will actually search:

- Include the client name, the deliverable type, and the date — these are the three most common search terms.
- Avoid generic words without context: `Report`, `Data`, `File`, `Document` alone mean nothing. Use `Campaign-Performance-Report`, not `Report`.
- Use consistent terminology — if you call it a `Brief` in one folder, do not call it a `Summary` in another.
- Spell out abbreviations in descriptions unless they are standard: `Q1` is fine; `CRL-APR-REV` is not.

### Searching Files in Apps Script

```javascript
// Search within a specific folder — scope the search to avoid full-Drive scans
function findInFolder(folder, name) {
  const files = folder.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}

// Search with a Drive API query string (most powerful)
// query format: field operator value [and/or ...]
function searchFiles(query) {
  return DriveApp.searchFiles(query);
}

// Find all Sheets for a client modified in the last 7 days
function getRecentClientSheets(clientName) {
  const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
  const query = [
    `mimeType = 'application/vnd.google-apps.spreadsheet'`,
    `name contains '${clientName}'`,
    `modifiedDate > '${cutoff}'`,
    `trashed = false`
  ].join(' and ');
  return DriveApp.searchFiles(query);
}

// Find by exact name — use only when you know there is exactly one match
function findFileByName(name) {
  const files = DriveApp.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}
```

**Important:** `DriveApp.getFilesByName()` and `DriveApp.getFoldersByName()` search across all of Drive, not within a parent. Always call them on a folder object (`parentFolder.getFilesByName()`) to scope the search correctly.

### File Descriptions as Metadata

Drive supports a description field on every file — use it for metadata that does not fit in the name:

```javascript
file.setDescription(
  `Generated by TAKScripts on ${formattedDate}. Client: ${clientName}. ` +
  `Source sheet: ${sourceSheetUrl}. Do not edit this file directly.`
);
```

---

## 11. Cleanup Patterns

### Duplicate Detection

```javascript
function findDuplicatesInFolder(folder) {
  const seen = {};
  const duplicates = [];
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (seen[name]) {
      duplicates.push({ name, id: file.getId(), url: file.getUrl() });
    } else {
      seen[name] = true;
    }
  }
  return duplicates;
}
```

### Empty Folder Cleanup

```javascript
function removeEmptyFolders(parent) {
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    removeEmptyFolders(folder);  // recurse first so children are cleaned before parent check
    const hasFiles = folder.getFiles().hasNext();
    const hasSubFolders = folder.getFolders().hasNext();
    if (!hasFiles && !hasSubFolders) {
      folder.setTrashed(true);
      Logger.log(`Trashed empty folder: ${folder.getName()}`);
    }
  }
}
```

### Stale File Identification

```javascript
function getStaleFiles(folder, daysOld) {
  const cutoff = new Date(Date.now() - daysOld * 24 * 60 * 60 * 1000);
  const stale = [];
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    if (file.getLastUpdated() < cutoff) {
      stale.push({
        name: file.getName(),
        lastUpdated: file.getLastUpdated(),
        id: file.getId(),
      });
    }
  }
  return stale;
}
```

### Cleanup Script Safety Pattern

Never auto-delete. Always run a dry-run first, log candidates, and require a second explicit call to execute:

```javascript
function cleanupCandidates(folderId, daysOld = 180, dryRun = true) {
  const folder = DriveApp.getFolderById(folderId);
  const stale = getStaleFiles(folder, daysOld);

  stale.forEach(f => {
    if (dryRun) {
      Logger.log(`[DRY RUN] Would archive: ${f.name} — last updated: ${f.lastUpdated}`);
    } else {
      const archiveFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties().getProperty('ARCHIVE_FOLDER_ID')
      );
      DriveApp.getFileById(f.id).moveTo(archiveFolder);
    }
  });

  Logger.log(`${stale.length} stale files found. dryRun=${dryRun}`);
}
```

---

## 12. Drive Storage Limits

### Quota Awareness

- Google Workspace accounts: storage depends on plan (30 GB per user at minimum, pooled in some tiers).
- Free accounts: 15 GB shared across Drive, Gmail, and Photos.
- **Google-native files (Sheets, Docs, Slides, Forms) do not count against quota.**
- **Uploaded files (PDF, XLSX, images, video) do count against quota.**

### Implications for Scripts

- Prefer creating native Google files when possible — no quota cost.
- When exporting to PDF for delivery, clean up the PDF after it has been sent if the native Sheet still exists. Do not accumulate PDFs for archival when the Sheet is the source of truth.
- Do not store large binaries (video, raw images) in Drive from scripts — use Cloud Storage for that.

### Logging File Size

```javascript
function logFileSize(file) {
  const sizeBytes = file.getSize();
  const sizeMB = (sizeBytes / 1_048_576).toFixed(2);
  Logger.log(`Created: ${file.getName()} — ${sizeMB} MB`);
  if (sizeBytes > 50_000_000) {
    Logger.log(`WARNING: ${file.getName()} exceeds 50 MB. Consider alternative storage.`);
  }
}
```

### Checking Quota via Drive API

Drive quota is not readable from `DriveApp`. Use the Drive REST API if you need to surface this:

```javascript
function getDriveQuota() {
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/about?fields=storageQuota',
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = JSON.parse(response.getContentText());
  // Returns: { limit, usage, usageInDrive, usageInDriveTrash }
  return data.storageQuota;
}
```

---

## 13. Apps Script Implementation Reference

### Core DriveApp Operations

```javascript
// Create a folder inside a parent
const parent = DriveApp.getFolderById('PARENT_ID');
const newFolder = parent.createFolder('2026-03-23_Acme-Corp_Q1-Reports');

// Create a plain text file inside a folder
const file = parent.createFile('notes.txt', 'Initial content', MimeType.PLAIN_TEXT);

// Move a file to a different folder (preferred over addFile/removeFile)
const file = DriveApp.getFileById('FILE_ID');
const destination = DriveApp.getFolderById('DEST_FOLDER_ID');
file.moveTo(destination);

// Move a folder to a different parent
const folder = DriveApp.getFolderById('FOLDER_ID');
folder.moveTo(destination);

// Copy a file to a destination
const copy = file.makeCopy('New Name', destination);

// Set folder color
folder.setColor(DriveApp.Color.GREEN);

// Set sharing
folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

// Add specific users
folder.addEditor('editor@example.com');
folder.addViewer('viewer@example.com');

// Set a file description
file.setDescription('Auto-generated by TAKScripts. Do not edit directly.');
```

### Idempotent Folder Creation

```javascript
// Returns the folder whether it existed already or was just created
function ensureFolder(parent, name) {
  const iterator = parent.getFoldersByName(name);
  return iterator.hasNext() ? iterator.next() : parent.createFolder(name);
}
```

### Full Client Folder Initialization

```javascript
function initializeClientFolder(clientName) {
  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty('ROOT_FOLDER_ID');
  const root = DriveApp.getFolderById(rootId);

  const safeName = sanitizeName(clientName);
  const clientFolder = ensureFolder(root, `${safeName} \u2014 TAKScripts`);

  // Create standard subfolders
  const subfolders = [
    '\uD83E\uDDFE Invoices',
    '\uD83D\uDCCA Reports',
    '\uD83D\uDCC4 Contracts',
    '\uD83D\uDDC4\uFE0F Archive',
  ];
  subfolders.forEach(name => ensureFolder(clientFolder, name));

  // Color-code active client root
  clientFolder.setColor(DriveApp.Color.GREEN);

  // Store the ID so future runs don't search by name
  props.setProperty(`CLIENT_FOLDER_${safeName}`, clientFolder.getId());

  return clientFolder;
}
```

### Searching Files by MIME Type

```javascript
// Get all Sheets in a folder
function getSheetsInFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  return folder.getFilesByType(MimeType.GOOGLE_SHEETS);
}

// Search across Drive with a query string
function findClientFiles(clientName) {
  const safe = clientName.replace(/'/g, "\\'");
  const query = `name contains '${safe}' and trashed = false`;
  const files = DriveApp.searchFiles(query);
  const results = [];
  while (files.hasNext()) {
    const f = files.next();
    results.push({ name: f.getName(), id: f.getId(), url: f.getUrl() });
  }
  return results;
}
```

---

## 14. What to Avoid

### Flat Structures

Dumping all files into one folder is the most common mistake. After ~20 files, retrieval is impractical and script iteration over folder contents becomes slow.

```
Bad:  My Drive/
      ├── Acme Report March.xlsx
      ├── acme report march FINAL.xlsx
      ├── BlueSky Invoice.pdf
      ├── Copy of Acme Report March.xlsx
      └── ... (200 more files)
```

### Spaces and Special Characters in Names

Spaces in file names break URL-encoded Drive links, make CLI tooling error-prone, and create inconsistency when names are used as keys elsewhere in the system.

```
Bad:  Acme Corp Q1 Report (Final!!).xlsx
Good: 2026-Q1_AcmeCorp_Campaign-Report_v2.xlsx
```

### Naming Mistakes That Break Sorting

- Using month names without a numeric prefix: `March Reports` sorts before `October Reports` alphabetically but after it chronologically
- Putting the day before the month: `23-03-2026` sorts by day, not year
- Using `v10` without zero-padding when you have more than 9 versions: `v2` sorts after `v10` — if that is ever possible, use `v02`, `v10`
- Starting names with articles: `The Monthly Report` pushes all these files together alphabetically

### Too-Deep Nesting

Beyond 4 levels, script code must chain multiple `getFolderById` calls and becomes brittle. It also hides context — if you need 6 levels to describe a file, the file naming is doing too little work.

### Permission Sprawl

- Never share individual files when a folder share suffices.
- Never grant `ANYONE` (public) access without an explicit product requirement.
- Never grant `Editor` when `Viewer` or `Commenter` is sufficient.
- Do not share files automatically during batch runs — separate sharing logic from creation logic so it can be reviewed and gated independently.

### Name-Based Lookup in Production

```javascript
// Bad: slow, fails silently when duplicates exist, fragile if someone renames the file
const file = DriveApp.getFilesByName('My Report').next();

// Good: always use stored IDs from Script Properties or a config sheet
const file = DriveApp.getFileById(storedFileId);
```

### Mutating Template Files

Never write to a template file directly. Always copy it first, then write to the copy. If a template file gets modified by a script, you lose your template — and probably won't notice until you need it.

### Trashing Instead of Archiving

`file.setTrashed(true)` gives a 30-day recovery window and then the file is gone. For business records, always move to an archive folder instead. Only call `setTrashed()` on confirmed junk (temp files, processing artifacts) that has no business value.

### Mixing Organizational Axes at the Same Level

Do not create a Level 2 where some folders are organized by date, some by client, and some by type — all as siblings. Pick one axis per level and apply it consistently. Mixed axes make the structure unpredictable for both humans and scripts.

### Over-Coloring

Applying all available colors to every folder level makes color meaningless. Reserve colors for status signals at one level (typically Level 2) only. If every folder is a different color, none of them are signaling anything.

---

## Quick Reference Cheatsheet

| Convention | Rule |
|---|---|
| Date format | `YYYY-MM-DD` prefix on all dated files and folders |
| Folder names | Title Case with spaces; use em dash for root brand separator |
| File segments | Separated by underscores; words within a segment separated by dashes |
| Client names | CamelCase, no spaces, normalized at intake, stored in a registry |
| File versions | `v1`, `v2` — never `_final` or `_new`; use `FINAL` only once |
| Hierarchy depth | Max 4 levels, default 3 |
| Admin folders | Prefix with `_` to sort to top and signal internal-only |
| Sharing | Folder-level preferred; `PRIVATE` as default; never auto-share in batch runs |
| Templates | Live in `_Templates/`, always copy before use, store IDs in Script Properties |
| Archive | Move to `_Archive/YYYY/`, never delete, color grey |
| Native formats | Use Sheets/Docs for editing; export PDF/CSV for delivery only |
| ID storage | Always store folder and file IDs in Script Properties, not names |
| Cleanup | Always run `dryRun = true` first; never auto-delete |
| Colors | Max 1 color scheme, Level 2 only, status-signal purpose only |
| Sanitization | Run all user-provided input through `sanitizeName()` before use in names |

---

*This document is part of the TAKScripts Polish Series. See also: POLISH_GUIDE.md, POLISH_SHEETS.md, POLISH_DOCS.md, POLISH_EMAIL.md, POLISH_SIDEBAR.md*
