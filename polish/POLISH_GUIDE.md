# POLISH_GUIDE — Master Polish Checklist & Workflow

The authoritative guide for the final polish stage before a TAKScripts Google Apps Script is marked approved.

---

## Table of Contents

1. [Overview](#1-overview)
2. [Polish Checklist](#2-polish-checklist)
3. [Per-Script Polish Notes](#3-per-script-polish-notes)
4. [Design Tokens](#4-design-tokens)
5. [Reference Files](#5-reference-files)
6. [Approval Criteria](#6-approval-criteria)

---

## 1. Overview

### What the Polish Stage Is

The polish stage is the final structured review pass before a script is marked **Approved** and published to customers. It is not about functionality — that was verified during testing. Polish is about presentation, consistency, and professionalism. A script that works correctly but looks rough will reflect poorly on the product and erode customer trust.

### When It Happens

Polish runs **after** the retesting pipeline has passed. The sequence is:

```
Development → QA Testing → Retesting (fixes verified) → POLISH → Approved
```

Do not begin polish on a script that still has open functional issues. Fix the code first, retest, then polish.

### What Polish Covers

- Visual consistency across Sheets dashboards, sidebars, dialogs, emails, and Docs output
- Language quality in user-facing text: menus, error messages, help sidebars, button labels
- Drive folder structure and file naming
- Adherence to the design token system (colors, fonts, sizes)
- No dead UI elements, broken layouts, or placeholder text remaining in production code

### Time Estimate

Budget **1–2 hours** per script for a thorough polish pass. Rushing this step is how scripts ship looking unfinished.

---

## 2. Polish Checklist

Work through each section that applies to the script being reviewed. Check every item before moving to the next section. Not all sections apply to every script — mark N/A where the script does not have that surface.

---

### 2.1 Sheets Dashboard

**Color Scheme**
- [ ] Header row uses the correct header background (`#1a73e8`) with white text
- [ ] Stat card rows use the correct stat card background (`#e8f0fe`)
- [ ] Status colors are from the approved token set only: success `#34a853`, warning `#fbbc04`, error `#ea4335`
- [ ] No off-palette colors have been introduced (no default Google red fills, no random grays)
- [ ] Alternating row bands use white and `#f8f9fa` (light gray), not any other shade
- [ ] Conditional formatting rules use the correct status colors and do not conflict with each other

**Row Heights**
- [ ] Data rows: 21px
- [ ] Header row: 30px
- [ ] Stat card rows: 60px
- [ ] No rows left at the default 21px when they should be a different height
- [ ] Row heights are set programmatically, not left to the sheet's default resize behavior

**Fonts**
- [ ] No explicit font is forced on Sheets cells (Sheets inherits the spreadsheet's default font — do not override unless intentional)
- [ ] Bold is used only on header rows, stat card labels, and column headers — not on body data
- [ ] Font sizes: 10–11pt for data rows, 12pt for column headers, 14–16pt for stat card values

**Column Widths**
- [ ] All columns have been explicitly sized — no columns left at default width
- [ ] Date columns: 100–110px
- [ ] Status columns: 90–110px
- [ ] Name/description columns: 180–250px
- [ ] Amount/count columns: 80–100px
- [ ] No column content is clipped or wrapping unexpectedly

**Stat Cards**
- [ ] Each stat card has a label (smaller, gray) and a value (larger, bold, colored)
- [ ] Stat card cells use `#e8f0fe` background fill
- [ ] Stat values use a semantic color where applicable (green for good, red for bad, blue for neutral counts)
- [ ] Stat cards are horizontally grouped at the top of the sheet, not scattered

**Header Row**
- [ ] Header row background: `#1a73e8`
- [ ] Header row text: white (`#ffffff`), bold
- [ ] Header row height: 30px
- [ ] Column label text matches the actual data in each column (no stale labels)
- [ ] Header row is frozen (row 1 or row 2 if there is a title row above it)

**Sheet Tabs**
- [ ] All tabs have human-readable names (no "Sheet1", "Sheet2")
- [ ] Dashboard tab is the first/leftmost tab
- [ ] Log/history tabs are to the right of the dashboard tab
- [ ] Tab colors are applied consistently if used at all

**General**
- [ ] No empty rows at the top above the header
- [ ] No placeholder or test data left in the sheet
- [ ] Gridlines are hidden on dashboard-style sheets (View > Gridlines off via script)
- [ ] Sheet is protected appropriately — user should not accidentally edit formula cells

---

### 2.2 Sidebar / Dialog UI

**Form Width**
- [ ] Sidebar width is set to `300px` (standard Apps Script sidebar width — do not override with a wider CSS value that causes horizontal scroll)
- [ ] Dialog width is appropriate for its content: settings dialogs 400–500px, confirmation dialogs 300–400px
- [ ] No content overflows horizontally at default zoom

**Save Confirmation**
- [ ] A visible success message appears after save — not just a button state change
- [ ] Success message uses green (`#34a853`) or green-tinted background
- [ ] Success message auto-dismisses after 3–5 seconds OR has a clear dismiss action
- [ ] Error messages use red (`#ea4335`) and explain what went wrong in plain language
- [ ] Save button returns to its default state after save (not stuck in "Saving..." state)

**Button States**
- [ ] Primary action button: `#1a73e8` background, white text, full width or right-aligned
- [ ] Destructive action button (delete, clear, reset): red (`#ea4335`) or outlined — never the same style as the primary button
- [ ] Disabled state is visually distinct (reduced opacity or gray background)
- [ ] Buttons have hover states (`:hover` CSS applied)
- [ ] No two buttons of equal visual weight are placed side by side with conflicting actions (e.g., "Run Now" next to "Clear All Data")

**Scrollability**
- [ ] Sidebar scrolls correctly when content exceeds viewport height
- [ ] No fixed-height containers that cut off content on smaller screens
- [ ] Sticky footer with action buttons works if implemented

**Typography**
- [ ] Font family: `Roboto, Arial, sans-serif` (in that order)
- [ ] Base font size: 13–14px for body text, 11–12px for helper text
- [ ] Section headings inside sidebars: 13–14px, bold, with a bottom border or margin separator
- [ ] No Times New Roman, Serif, or monospace fonts in user-facing sidebar UI

**Layout**
- [ ] Labels are left-aligned above their inputs (not inline to the left)
- [ ] Input fields are full-width within the sidebar content area
- [ ] Consistent vertical spacing between form groups: 12–16px gap
- [ ] Sidebar has 12–16px padding on all sides — no content flush against the edge

---

### 2.3 Email Output

**Subject Line**
- [ ] Subject line is specific and informative — a recipient can understand the email's purpose from the subject alone
- [ ] Subject line uses Title Case or Sentence case consistently — not ALL CAPS
- [ ] Subject line includes dynamic context where appropriate (e.g., client name, date, count)
- [ ] Subject line is under 70 characters to avoid truncation in most email clients

**HTML Structure**
- [ ] Email uses a single-column layout (max-width 600px, centered)
- [ ] All styles are inline — no `<style>` block, no external CSS (Gmail strips `<head>` styles)
- [ ] Font stack: `Arial, Helvetica, sans-serif` (web-safe fonts only — no Roboto, no Google Fonts)
- [ ] Table-based layout if complex structure is needed — no CSS flexbox or grid in email
- [ ] Background color is white or very light gray (`#f8f9fa`) — no dark mode-breaking backgrounds
- [ ] Email renders correctly at 375px width (mobile) and 600px width (desktop)

**Status Colors in Email**
- [ ] Success / positive states: `#34a853` (green)
- [ ] Warning / attention states: `#fbbc04` (yellow/amber)
- [ ] Error / negative states: `#ea4335` (red)
- [ ] Neutral / informational: `#1a73e8` (blue)
- [ ] Color is never the only indicator — color is paired with text ("Status: Overdue", not just a red cell)

**Footer**
- [ ] Footer is present on every automated email
- [ ] Footer includes the script name and/or "Sent by TAKScripts"
- [ ] Footer includes an unsubscribe note or settings reference where appropriate (for digest-style recurring emails)
- [ ] Footer text is smaller (12px) and gray (`#5f6368`)
- [ ] Footer is separated from body content with a top border or extra spacing

**Content**
- [ ] No lorem ipsum or placeholder text
- [ ] No raw JSON, raw array output, or developer-facing error text in the email body
- [ ] Numbers are formatted with commas and appropriate decimal places (not `1234.5` — use `$1,234.50`)
- [ ] Dates are human-readable (not ISO strings like `2026-03-23T00:00:00.000Z`)

---

### 2.4 Docs Output

**Page Margins**
- [ ] Standard business document: 1.0" top/bottom, 1.25" left/right
- [ ] Invoice (compact): 0.75" top/bottom, 1.0" left/right
- [ ] Contract/legal: 1.0" top/bottom, 1.5" left/right
- [ ] Margins are set programmatically — not left at the Docs default

**Fonts**
- [ ] Body text: 10–11pt, default Docs font (Arial or the doc's default)
- [ ] Section headings: 12–14pt, bold
- [ ] Cover page title: 20–24pt, bold, centered
- [ ] No font mixing within the same document — one body font, one heading font
- [ ] No Comic Sans, no Times New Roman in client-facing docs (use Arial, Roboto, or Georgia)

**Table Styles**
- [ ] Header row of every table has a colored background (`#1a73e8` for primary tables or `#e8eaed` for secondary)
- [ ] Header row text is bold and white (or dark on light headers)
- [ ] Table borders are present and consistent — no half-bordered tables
- [ ] Column widths are explicitly set, not left to auto-fit
- [ ] Alternating row bands applied where the table has more than 5 data rows
- [ ] No empty rows inside tables (no spacer rows used for padding — use cell padding instead)

**Cover Page**
- [ ] Present for proposals and contracts — not required for invoices
- [ ] Includes: document title, client name, date, preparer name or company
- [ ] Cover page is visually distinct from body pages (centered content, larger title)
- [ ] Page break follows the cover page so body starts on a new page

**General**
- [ ] Page numbers are in the footer for multi-page documents
- [ ] Document title in Docs matches the filename in Drive
- [ ] No tracked changes or comments left in the document
- [ ] No developer debug text or template placeholders (`{{CLIENT_NAME}}`, `[INSERT HERE]`) remaining

---

### 2.5 Drive Organization

**Folder Naming**
- [ ] All client-facing folders use Title Case with spaces (e.g., `Invoice Archive`, not `invoice_archive`)
- [ ] Root folder follows the pattern: `[ClientName] — TAKScripts` (em dash `\u2014`, not a hyphen)
- [ ] Emoji prefixes are used consistently — if any folder has an emoji, all sibling folders at that level should too
- [ ] No folder names longer than 50 characters
- [ ] No illegal characters in folder names: no `/`, `:`, `*`, `?`, `"`, `|`

**File Naming**
- [ ] Pattern: `[YYYY-MM-DD]_[Type]_[Client]_[Description].[ext]`
- [ ] Dates use ISO 8601 format (`2026-03-23`, not `March 23` or `03-23-2026`)
- [ ] Segments separated by underscores, words within segments separated by dashes
- [ ] No spaces in file names (replaced by dashes or CamelCase)
- [ ] File name length under 100 characters

**Permissions**
- [ ] Client-shared folders contain only client-appropriate content (no internal notes, no other clients' data)
- [ ] Script does not share files publicly ("Anyone with the link") unless that is explicitly the feature
- [ ] Sharing is applied at the folder level, not the individual file level, where possible

**Structure**
- [ ] Folder depth does not exceed 3 levels (4 maximum with strong justification)
- [ ] Active and archive content are separated — no mixing of current and historical files in the same folder
- [ ] No empty folders left after a run (create folders only when content exists to place in them)

---

### 2.6 Menu Items

**Naming Conventions**
- [ ] Menu items use Title Case: "Run Now", "Open Settings", "View Log" — not "run now" or "RUN NOW"
- [ ] Menu item names are action-oriented (verb + noun): "Generate Invoice", "Send Digest", "Clear Log"
- [ ] Menu item names are specific enough to be self-explanatory without needing documentation

**Grouping**
- [ ] Primary actions are at the top of the menu
- [ ] Destructive or dangerous actions (clear, delete, reset) are at the bottom
- [ ] Settings and Help are the last items in the menu

**Separators**
- [ ] Separators (`.addSeparator()`) are used to group related actions — not scattered randomly
- [ ] No more than 3–4 groups of items in a single script menu
- [ ] Help and About items are separated from functional items with a separator

**General**
- [ ] Menu name matches the product name as listed in the marketplace listing
- [ ] No "Test", "Debug", or "Dev Only" items left in production menus
- [ ] Menu items that trigger long operations show a toast or sidebar — not silent execution

---

### 2.7 Error Messages

**User-Facing Messages**
- [ ] Error messages are written in plain language — no stack traces, no JavaScript error objects
- [ ] Error messages explain what went wrong AND what the user should do next
- [ ] Error messages use the word "could not" or "unable to" rather than "Error:" as an opener when possible
- [ ] No error message contains internal variable names, function names, or line numbers
- [ ] Validation errors on sidebar forms highlight the specific field that failed — not just a generic "Something went wrong"

**Logger.log Quality**
- [ ] Every function that modifies data has at least one `Logger.log()` call confirming what it did
- [ ] Log messages are descriptive: `Logger.log('Created invoice folder: ' + folderName)` not `Logger.log('done')`
- [ ] Error catch blocks always log the error: `Logger.log('Error in generateInvoice: ' + e.message)`
- [ ] No sensitive data (passwords, full email bodies, personal info) in Logger.log output
- [ ] Log messages use consistent prefixes so they are easy to scan: `[INFO]`, `[WARN]`, `[ERROR]`

---

### 2.8 Help Sidebar

**Completeness**
- [ ] Every menu item has a corresponding entry in the Help sidebar explaining what it does
- [ ] Every input field in the Settings sidebar is explained in Help
- [ ] Known limitations or "gotchas" are documented in Help (e.g., "This script only processes the first 500 rows")

**Accuracy**
- [ ] Help content matches the current version of the script — no outdated instructions
- [ ] Screenshots or diagrams referenced in Help actually exist and are current
- [ ] Step-by-step instructions have been walked through manually to confirm they are correct

**Formatting**
- [ ] Sections are clearly headed and separated
- [ ] Numbered lists for sequential steps, bullet lists for non-sequential items
- [ ] No walls of unbroken text — paragraphs are short (3–5 sentences max)
- [ ] Link text is descriptive ("View setup guide") not generic ("Click here")

---

## 3. Per-Script Polish Notes

Key surfaces to check during polish for each script. Use these alongside the full checklist above — these notes call out the highest-risk areas specific to each script.

---

### 3.1 auto-invoice-generator

| Surface | Key Polish Points |
|---------|------------------|
| Docs output | Cover page present with client name, date, invoice number. Table borders consistent. Line items formatted as currency. Total row bold and clearly separated. |
| Email | Subject includes invoice number and client name. Attachment confirmed. Amount formatted as `$X,XXX.XX`. Payment instructions clear. |
| Drive folders | Root folder uses em dash pattern. Invoice subfolder named `🧾 Invoices`. Files named `[Date]_Invoice_[Client]_[InvoiceNumber].pdf`. |

---

### 3.2 contract-proposal-generator

| Surface | Key Polish Points |
|---------|------------------|
| Docs output | Cover page with client name, proposal date, and preparer. Section headings consistent size and style. Signature block has adequate white space below signature lines. |
| Drive folders | Separate subfolders for Proposals and Contracts if both types are generated. Version suffix on file name if revisions are possible. |
| Sidebar | Client name and project title fields validated before generation. Currency fields formatted on input (no raw numbers). Clear success message with a link to the generated doc. |

---

### 3.3 email-invoice-detector

| Surface | Key Polish Points |
|---------|------------------|
| Sheets dashboard | Status column uses conditional formatting: detected (green), pending (yellow), error (red). Invoice count stat card at top. Date detected column formatted as readable date, not timestamp. |
| Email alerts | Alert subject specific: "Invoice Detected: [Sender] — [Amount]". No raw Gmail message IDs in the email body. |

---

### 3.4 follow-up-nudger

| Surface | Key Polish Points |
|---------|------------------|
| Sheets dashboard | Days overdue column uses conditional formatting: 1–3 days (yellow), 4–7 days (orange), 8+ days (red). Last contacted date formatted as `MMM d, yyyy`. |
| Email digest | Digest groups follow-ups by urgency level. Each follow-up entry shows: contact name, company, days overdue, last contact date. Footer includes link to the dashboard sheet. |
| Sidebar | Follow-up threshold (days) input has sensible default (3) pre-filled. Save confirmation visible and auto-dismisses. |

---

### 3.5 inventory-low-stock-alert

| Surface | Key Polish Points |
|---------|------------------|
| Sheets dashboard | Current stock column: red text when below threshold, green when at or above. Threshold column is editable but highlighted to show it is a config column. Last updated timestamp in stat card area. |
| Email alerts | Subject: "Low Stock Alert: [N] items below threshold". Item list in email formatted as a table, not a raw list. Reorder quantity suggestion included if that feature is active. |
| Sidebar | Threshold default pre-filled. Email recipient field shows the currently saved address. Run Now button triggers a confirmation toast. |

---

### 3.6 vip-priority-alert

| Surface | Key Polish Points |
|---------|------------------|
| Sheets dashboard | VIP list is editable directly in the sheet with clear column headers. Priority level column uses color coding: High (red), Medium (yellow), Low (green). Last alert sent column present and formatted. |
| Email alerts | Alert subject identifies the VIP by name. Alert email is brief — 3–5 lines max. Clearly distinguishable from other automated emails (clear "VIP Alert" branding in header). |
| Sidebar | VIP management UI is intuitive — add/remove from list without needing to edit the sheet directly if a sidebar UI is provided. |

---

### 3.7 smart-ooo-auto-reply

| Surface | Key Polish Points |
|---------|------------------|
| Email reply templates | Reply template is professional and free of placeholder text. Return date is dynamically inserted and formatted as a human-readable date (`Monday, March 30`). VIP contacts get a different, more detailed reply if that feature is active. |
| Sidebar | Date range picker (start/end dates) clearly labeled. On/Off toggle state is visually obvious. Current status ("Auto-reply is ON until March 30") shown prominently at the top of the sidebar. |

---

### 3.8 bulk-email-unsubscriber

| Surface | Key Polish Points |
|---------|------------------|
| Sheets dashboard | Unsubscribed count and processed count as stat cards. List of unsubscribed senders formatted as a clean table with timestamp. Status column: Unsubscribed (green), Failed (red), Skipped (gray). |
| Sidebar | Run scope clearly explained (last N days, or specific label). Dry run mode visually distinct from live run mode — different button color or a warning banner. Confirmation dialog before executing a live run. |

---

### 3.9 auto-file-organizer

| Surface | Key Polish Points |
|---------|------------------|
| Drive folders | Source and destination folder IDs validated before running. Destination folder structure matches the naming and hierarchy standards. No files moved to root of Drive. |
| Sheets log | Log sheet records: file name, source folder, destination folder, timestamp, and status. Log rows use status color coding. Newest entries at the top (or log is sorted by timestamp descending). |
| Sidebar | Folder picker or folder ID input clearly labeled. Preview/dry run mode shows what would be moved before executing. |

---

### 3.10 attachment-auto-saver

| Surface | Key Polish Points |
|---------|------------------|
| Drive folders | Saved attachments land in clearly named subfolders (e.g., by sender domain or by month). File names use the standard pattern: `[Date]_[OriginalName]` or `[Date]_[Sender]_[OriginalName]`. |
| Sheets log | Log captures: attachment name, sender email, Gmail message date, save destination, and save timestamp. Log is filterable by sender or date. |
| Sidebar | Filter rules (sender, subject, file type) are clearly labeled with examples. Saved configuration is displayed back to the user after saving. |

---

### 3.11 data-cleanup-wizard

| Surface | Key Polish Points |
|---------|------------------|
| Sheets output | Cleaned cells are highlighted briefly or a summary row added showing what was changed. Before/after counts shown as stat cards: rows processed, cells modified, duplicates removed. |
| Sidebar | Cleanup options are grouped logically: Format fixes, Duplicate removal, Blank row removal. Destructive operations (delete rows) require a checkbox confirmation before the Run button becomes active. Preview of changes shown before committing where feasible. |

---

## 4. Design Tokens

These are the canonical values for all TAKScripts Google Apps Scripts. Use these and only these. Do not introduce new colors, fonts, or size values without updating this token table.

### 4.1 Colors

| Token | Hex | Usage |
|-------|-----|-------|
| `color-primary` | `#1a73e8` | Dashboard header bg, sidebar primary button, link color |
| `color-primary-text` | `#ffffff` | Text on primary color backgrounds |
| `color-stat-bg` | `#e8f0fe` | Stat card background in Sheets |
| `color-success` | `#34a853` | Success status, positive metrics, confirmed actions |
| `color-warning` | `#fbbc04` | Warning status, pending items, attention required |
| `color-error` | `#ea4335` | Error status, overdue items, failed operations |
| `color-neutral-text` | `#3c4043` | Primary body text in sidebars and emails |
| `color-secondary-text` | `#5f6368` | Helper text, labels, footer text |
| `color-border` | `#dadce0` | Input borders, table borders, dividers |
| `color-row-alt` | `#f8f9fa` | Alternating row band in Sheets |
| `color-surface` | `#ffffff` | Background for sidebars, dialogs, email bodies |
| `color-destructive` | `#ea4335` | Destructive action buttons |

### 4.2 Fonts

| Context | Font Stack | Notes |
|---------|-----------|-------|
| Sidebar / Dialog UI | `Roboto, Arial, sans-serif` | Load Roboto from Google Fonts CDN in sidebar HTML |
| Email HTML | `Arial, Helvetica, sans-serif` | Web-safe only — no Google Fonts in email |
| Sheets cells | (inherit spreadsheet default) | Do not force a font family in Sheets; let it inherit |
| Docs body | `Arial` or document default | Match the document's established font — do not mix |

### 4.3 Sizes — Sheets Row Heights

| Row Type | Height |
|----------|--------|
| Data row | 21px |
| Header / column label row | 30px |
| Stat card row | 60px |
| Section divider row | 8px |

### 4.4 Sizes — Typography

| Element | Size | Weight |
|---------|------|--------|
| Sidebar body text | 13–14px | Regular |
| Sidebar helper / label text | 11–12px | Regular |
| Sidebar section heading | 13–14px | Bold |
| Email body text | 14px | Regular |
| Email section heading | 16–18px | Bold |
| Email footer text | 12px | Regular |
| Sheets data cell | 10–11pt | Regular |
| Sheets column header | 12pt | Bold |
| Sheets stat card value | 14–16pt | Bold |
| Docs body | 10–11pt | Regular |
| Docs section heading | 12–14pt | Bold |
| Docs cover title | 20–24pt | Bold |

### 4.5 Sizes — UI Components

| Component | Value |
|-----------|-------|
| Sidebar width | 300px (standard) |
| Sidebar padding | 12–16px all sides |
| Form group vertical gap | 12–16px |
| Primary button height | 36px |
| Input field height | 32–36px |
| Email max-width | 600px |
| Dialog width (settings) | 400–500px |
| Dialog width (confirmation) | 300–400px |

### 4.6 Status Color Application Rules

Status colors must always be paired with text labels — never used as the sole indicator.

| Status | Background | Text | Text Label |
|--------|-----------|------|------------|
| Success / Complete | `#e6f4ea` | `#137333` | "Complete", "Sent", "Saved" |
| Warning / Pending | `#fef7e0` | `#b45309` | "Pending", "Awaiting", "Due Soon" |
| Error / Failed | `#fce8e6` | `#c5221f` | "Failed", "Error", "Overdue" |
| Neutral / Active | `#e8f0fe` | `#1a73e8` | "Active", "Running", "Processing" |

---

## 5. Reference Files

Each reference file covers one output surface in depth — design rationale, implementation patterns, and what to avoid. Consult the relevant file during the polish pass for that surface.

| File | Covers |
|------|--------|
| `POLISH_SHEETS.md` | Color system, row heights, column widths, header rows, stat cards, conditional formatting, sheet tabs, Apps Script implementation |
| `POLISH_DOCS.md` | Page setup, margins, typography system, color system, document structure, headings, tables, cover pages, proposals, contracts, invoices |
| `POLISH_EMAIL.md` | HTML email structure, inline CSS patterns, subject lines, status color use, footer standards, mobile rendering |
| `POLISH_SIDEBAR.md` | Form layout, button states, typography, scrollability, save/error confirmation patterns, accessibility basics |
| `POLISH_DRIVE.md` | Folder naming conventions, file naming conventions, date formats, folder hierarchy, archive structure, emoji prefixes, client-shared folder design |

> Note: `POLISH_EMAIL.md` and `POLISH_SIDEBAR.md` should be created if not yet present. The patterns in Section 2.3 and 2.2 of this guide serve as the interim reference until those files exist.

---

## 6. Approval Criteria

A script may only be marked **Approved** when all of the following are true.

### 6.1 Functional Requirements (Pre-Polish Gate)

- [ ] All retesting runs passed with no open bugs
- [ ] No console errors or unhandled exceptions during a normal run
- [ ] The script completes within Google Apps Script's 6-minute execution limit for typical data volumes
- [ ] Trigger setup (if applicable) has been tested and confirmed working

### 6.2 Visual Requirements

- [ ] All applicable sections of the Polish Checklist (Section 2) are checked off
- [ ] No design token violations — every color, font, and size is from the token table in Section 4
- [ ] No placeholder text, test data, or debug output visible in any user-facing surface
- [ ] At least one complete end-to-end run reviewed visually by the developer (not just logged output)

### 6.3 Text & Language Requirements

- [ ] All user-facing text (menus, buttons, error messages, emails, help sidebar) has been read aloud or reviewed for grammar and clarity
- [ ] No developer jargon in user-facing messages
- [ ] Script name in the menu matches the product name in the listing exactly

### 6.4 Drive & File Requirements

- [ ] Folder naming follows the conventions in `POLISH_DRIVE.md`
- [ ] File naming follows the `[Date]_[Type]_[Client]_[Description].[ext]` pattern
- [ ] No files are created in unexpected locations (root of Drive, wrong parent folder)

### 6.5 Final Sign-Off Checklist

Before changing status to Approved, confirm:

- [ ] Polish checklist completed and all items checked or marked N/A
- [ ] Per-script notes for this script reviewed (Section 3)
- [ ] One final full run executed after all polish edits — no regressions introduced during polish
- [ ] Script version number or last-edited date updated in the script's comment header
- [ ] Listing description and feature bullets reviewed for accuracy against the final script behavior

---

*This document is the master reference for the polish stage. See also: POLISH_SHEETS.md, POLISH_DOCS.md, POLISH_EMAIL.md, POLISH_SIDEBAR.md, POLISH_DRIVE.md*
