# Google Docs Design Reference — Polish Guide for Apps Script Documents

A comprehensive reference for developers building Google Apps Scripts that generate professional Google Docs (proposals, contracts, invoices, reports). Follow these standards to produce documents that look polished, consistent, and client-ready.

---

## Table of Contents

1. [Page Setup](#1-page-setup)
2. [Typography](#2-typography)
3. [Color System](#3-color-system)
4. [Headers and Footers](#4-headers-and-footers)
5. [Cover Pages](#5-cover-pages)
6. [Tables](#6-tables)
7. [Paragraph Styles](#7-paragraph-styles)
8. [Lists](#8-lists)
9. [Images](#9-images)
10. [Document Structure](#10-document-structure)
11. [Invoice / Proposal / Contract Layouts](#11-invoice--proposal--contract-layouts)
12. [What to Avoid](#12-what-to-avoid)
13. [Apps Script Implementation Notes](#13-apps-script-implementation-notes)

---

## 1. Page Setup

### Standard vs. Narrow Margins

| Use Case | Top | Bottom | Left | Right | Notes |
|---|---|---|---|---|---|
| Proposals / Contracts | 1 in | 1 in | 1.25 in | 1.25 in | Comfortable reading, professional feel |
| Invoices | 0.75 in | 0.75 in | 1 in | 1 in | Fits more line items without crowding |
| Reports (dense data) | 0.75 in | 0.75 in | 0.75 in | 0.75 in | Narrow; use only when data density demands it |
| Cover Pages | 1.5 in | 1.5 in | 1.5 in | 1.5 in | Generous whitespace reads as premium |

**Rule of thumb:** Default to 1 in margins. Only go narrow if you have a specific space problem. Wider margins on cover pages make the document feel intentional, not squeezed.

**Left margin note for contracts:** Printed contracts are frequently three-hole punched or bound on the left. A 1.25–1.5 in left margin protects text from being obscured by the binding.

### Page Size

Always use **Letter (8.5 × 11 in)** for US clients. Use **A4 (8.27 × 11.69 in)** only if the client is outside North America. Mixing sizes in a multi-section document is never appropriate.

### Orientation

- **Portrait** — default for all document types.
- **Landscape** — acceptable only for wide financial tables or architectural drawings embedded as a separate section. Never use landscape for an entire proposal or contract.

### Apps Script: Page Setup

```javascript
// Google Docs API uses points internally — 1 inch = 72 pt
function setPageSetup(doc) {
  const body = doc.getBody();

  // Standard margins for proposals/contracts
  body.setMarginTop(72);      // 1.0 in
  body.setMarginBottom(72);   // 1.0 in
  body.setMarginLeft(90);     // 1.25 in
  body.setMarginRight(90);    // 1.25 in

  // Narrow margins for invoices:
  // body.setMarginTop(54);   // 0.75 in
  // body.setMarginLeft(72);
  // body.setMarginRight(72);
}

// Page size requires the Docs REST API (DocumentApp does not expose setPageWidth)
function setPageSizeViaRestApi(docId) {
  const requests = [{
    updateDocumentStyle: {
      documentStyle: {
        pageSize: {
          width:  { magnitude: 612, unit: 'PT' },  // Letter: 8.5 in
          height: { magnitude: 792, unit: 'PT' }   // Letter: 11 in
        },
        marginTop:    { magnitude: 72, unit: 'PT' },
        marginBottom: { magnitude: 72, unit: 'PT' },
        marginLeft:   { magnitude: 90, unit: 'PT' },
        marginRight:  { magnitude: 90, unit: 'PT' }
      },
      fields: 'pageSize,marginTop,marginBottom,marginLeft,marginRight'
    }
  }];

  Docs.Documents.batchUpdate({ requests }, docId);
}
```

---

## 2. Typography

### Recommended Font Pairings

Google Docs is limited to fonts available in the Google Fonts library. These pairings produce professional, readable documents:

#### Pairing A — Clean and Modern (default recommendation)

| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Montserrat** | Bold (700) | See heading scale below |
| Body | **Open Sans** | Regular (400) | 11 pt |
| Body emphasis | **Open Sans** | SemiBold (600) | 11 pt |
| Monospace / codes | **Roboto Mono** | Regular | 10 pt |

#### Pairing B — Warm and Professional (law firms, consultants)

| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Roboto** | Bold (700) | See heading scale below |
| Body | **Georgia** | Regular | 11 pt |
| Body emphasis | **Georgia** | Bold | 11 pt |

#### Pairing C — Minimal / Tech

| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Lato** | Bold | See heading scale below |
| Body | **Lato** | Regular | 11 pt |

**Never mix more than two typefaces in a single document.** Headings + body = 2 fonts maximum. A monospace font for code snippets or reference numbers is an acceptable third but must be used sparingly and consistently.

### Heading Size Scale

| Level | Style Name in Docs | Font Size | Weight | Case |
|---|---|---|---|---|
| Document Title | Title | 26–28 pt | Bold | Title Case |
| Subtitle | Subtitle | 14–16 pt | Regular or Light | Title Case |
| H1 — major sections | Heading 1 | 18 pt | Bold | Title Case |
| H2 — subsections | Heading 2 | 14 pt | Bold | Title Case |
| H3 — sub-subsections | Heading 3 | 12 pt | Bold | Sentence case |
| H4 — rarely needed | Heading 4 | 11 pt | Bold | Sentence case |

**Scale down consistently.** There must be a visible size difference between each level. If H1 and H2 look the same, one of them is wrong. If you find yourself needing H4 frequently, the document's content hierarchy needs restructuring.

### Body Text

- **Size:** 11 pt for most documents. Use 10.5 pt only if fitting to a page is critical (e.g., a one-page invoice).
- **Line spacing:** 1.15 is the Google Docs default and acceptable. For proposals, use **1.2** for a slightly more airy feel.
- **Paragraph spacing:** Add **6–8 pt after** each paragraph. This replaces blank lines between paragraphs — never use empty paragraph breaks to create vertical space.
- **Color:** `#1A1A1A` (near-black). Pure black (`#000000`) is technically fine but feels harsh on-screen.

### Line Length

At 8.5 in wide with 1.25 in margins, your text column is 6 in — approximately 65–75 characters at 11 pt. This is the ideal range for comfortable reading. Do not go narrower (by increasing margins past 1.5 in) without good reason, and do not go wider.

---

## 3. Color System

### Color Roles

Define a small, intentional palette before writing any script. Every color in the document must serve a named role.

```
PRIMARY      — H1 headings, cover page elements, table headers
SECONDARY    — H2 / H3 headings, section labels
ACCENT       — Totals row, callout boxes, CTA elements, horizontal dividers
TABLE_HEADER — Table header row backgrounds (usually same as PRIMARY)
TABLE_ALT    — Alternating row tint (subtle)
TEXT_DARK    — Primary body text
TEXT_MUTED   — Captions, footnotes, secondary labels, header/footer text
TABLE_BORDER — Table borders and dividers
WHITE        — Text on dark backgrounds
```

### Example Palette — Professional Blue

```javascript
const COLORS = {
  PRIMARY:      '#1B3A6B',  // Deep navy — H1, cover, table headers
  SECONDARY:    '#2E6DA4',  // Medium blue — H2, H3, accents
  ACCENT:       '#E87722',  // Warm orange — totals, CTA, highlights
  TABLE_HEADER: '#1B3A6B',  // Same as PRIMARY
  TABLE_ALT:    '#EEF3F9',  // Very light blue tint — alternating rows
  TABLE_BORDER: '#C5D5E8',  // Soft blue-gray
  TEXT_DARK:    '#1A1A1A',  // Near-black body text
  TEXT_MUTED:   '#6B7280',  // Gray — captions, dates, footer text
  DIVIDER:      '#D1D5DB',  // Light gray horizontal rules
  WHITE:        '#FFFFFF',  // Text on colored backgrounds
};
```

### Example Palette — Professional Green (consulting / sustainability)

```javascript
const COLORS = {
  PRIMARY:      '#1A4731',  // Deep forest green
  SECONDARY:    '#2D7A4F',  // Mid green
  ACCENT:       '#F59E0B',  // Amber
  TABLE_HEADER: '#1A4731',
  TABLE_ALT:    '#EDF7F1',  // Light green tint
  TABLE_BORDER: '#B7D9C6',
  TEXT_DARK:    '#1A1A1A',
  TEXT_MUTED:   '#6B7280',
  DIVIDER:      '#D1D5DB',
  WHITE:        '#FFFFFF',
};
```

### Example Palette — Neutral / Corporate Gray

```javascript
const COLORS = {
  PRIMARY:      '#1F2937',  // Charcoal
  SECONDARY:    '#374151',  // Dark gray
  ACCENT:       '#3B82F6',  // Blue accent
  TABLE_HEADER: '#1F2937',
  TABLE_ALT:    '#F3F4F6',  // Very light gray
  TABLE_BORDER: '#D1D5DB',
  TEXT_DARK:    '#111827',
  TEXT_MUTED:   '#6B7280',
  DIVIDER:      '#E5E7EB',
  WHITE:        '#FFFFFF',
};
```

### Color Rules

- **Text on a dark background** must be `#FFFFFF` or near-white. Never use `#1A1A1A` text on `#1B3A6B` background.
- **Alternating row tint** should be barely perceptible — a contrast ratio below 1.15:1 against white. It is a whisper, not a stripe.
- **Accent color** appears sparingly — totals row, one callout, one horizontal divider. If it appears on 8+ elements, it has lost its function.
- **Never use red** for emphasis in financial documents. Red signals a negative value or error. Use the accent color or bold text instead.
- **Heading colors** should be `PRIMARY` for H1, `SECONDARY` for H2/H3. Body text is always `TEXT_DARK`.

---

## 4. Headers and Footers

### Header Content by Document Type

| Document Type | Left | Center | Right |
|---|---|---|---|
| Proposal | Company logo or name | — | Document title |
| Contract | Company name (text) | — | "CONFIDENTIAL" |
| Invoice | Company logo | — | Invoice number + date |
| Report | Report title | — | Company name |

**Logo in header:** Left-aligned is standard. Keep the repeating header logo small — 0.4–0.6 in tall, proportional width. A logo that dominates the header on every page steals attention from content.

**Suppress the header on the cover page** — the cover page has its own branding treatment. Use a different first-page header (via the document's header settings) or leave the cover page header blank.

### Footer Content

| Left | Center | Right |
|---|---|---|
| Company address or tagline | — | Page X of Y |
| Confidentiality notice (contracts) | — | Page X of Y |
| Document date | — | Page X of Y |

**Font size in headers/footers:** 8–9 pt. Color: `TEXT_MUTED` (`#6B7280`). Headers and footers are structural, not readable content — they must recede visually.

### Apps Script: Headers and Footers

```javascript
function addHeaderFooter(doc, docTitle) {
  // --- HEADER ---
  const header = doc.addHeader();

  // Use a borderless two-cell table for left + right alignment
  const headerTable = header.appendTable([[' ', docTitle]]);
  headerTable.setBorderWidth(0);

  const leftCell = headerTable.getCell(0, 0);
  leftCell.setText('Your Company Name');
  leftCell.editAsText().setFontSize(9).setForegroundColor('#6B7280');

  const rightCell = headerTable.getCell(0, 1);
  rightCell.setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rightCell.editAsText().setFontSize(9).setForegroundColor('#6B7280');

  // --- FOOTER ---
  const footer = doc.addFooter();
  const footerPara = footer.appendParagraph('');
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  footerPara.appendText('Page ');
  footerPara.editAsText().setFontSize(8).setForegroundColor('#6B7280');
  // Note: live PAGE / NUMPAGES fields require the Docs REST API (see below)
}

// Live page numbers via REST API batchUpdate
function insertPageNumberField(docId, segmentId, isFooter) {
  // Requires Advanced Google Services: Docs API enabled
  const requests = [{
    insertPageBreak: { /* ... position object ... */ }
  }];
  // Use Docs.Documents.batchUpdate with insertPageNumber request type
  // This is the only way to get dynamic page numbers via script
}
```

**Important limitation:** `DocumentApp` cannot insert live `PAGE` or `NUMPAGES` fields. For dynamic page numbers, use the Docs REST API `insertPageNumber` request in a `batchUpdate` call, or insert a static placeholder and instruct the user to update it.

---

## 5. Cover Pages

### Cover Page Layout

A cover page communicates who you are, what this document is, and creates a strong first impression. Use generous whitespace.

```
┌─────────────────────────────────┐
│  [Logo — top left, 1.5 in max]  │
│                                 │
│  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━  │  ← Primary-colored accent bar
│                                 │
│  DOCUMENT TITLE                 │  ← 28 pt, bold, PRIMARY color
│  Document Subtitle              │  ← 15 pt, SECONDARY color
│                                 │
│  Prepared for:                  │  ← 9 pt, TEXT_MUTED
│  Client Name                    │  ← 12 pt, TEXT_DARK
│  Client Company                 │
│                                 │
│  Prepared by:                   │  ← 9 pt, TEXT_MUTED
│  Your Name                      │  ← 12 pt, TEXT_DARK
│  Your Company                   │
│                                 │
│  Date: March 23, 2026           │  ← 10 pt, TEXT_MUTED
│                                 │
└─────────────────────────────────┘
```

### Cover Page Design Rules

- **One accent element** — a full-width colored bar in `PRIMARY` (~0.12–0.15 in thick) is clean and authoritative.
- **Title at 26–28 pt**, bold, `PRIMARY` color. Never all-caps for the full title.
- **Subtitle at 14–16 pt**, regular weight, `SECONDARY` or `TEXT_MUTED`.
- **"Prepared for/by" labels** at 9 pt, `TEXT_MUTED`. Client and company names at 12 pt, `TEXT_DARK`, not bold.
- **Date** at 9–10 pt, `TEXT_MUTED`.
- **No page number on the cover page.** Start page numbering from page 2 (the TOC or first content page).
- Always end with an explicit page break before the next section.

### Apps Script: Cover Page

```javascript
function buildCoverPage(body, data) {
  // Company name (logo placeholder — replace with image blob if available)
  const companyPara = body.appendParagraph(data.companyName);
  companyPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  companyPara.editAsText()
    .setFontFamily('Montserrat')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);

  // Spacer — cover pages are the one place where 2-3 blank paragraphs are acceptable
  body.appendParagraph('');
  body.appendParagraph('');

  // Accent bar (simulated with a single-cell, single-row table)
  const accentTable = body.appendTable([['']]);
  accentTable.setBorderWidth(0);
  const accentCell = accentTable.getCell(0, 0);
  accentCell.setBackgroundColor(COLORS.PRIMARY);
  accentTable.getRow(0).setMinimumHeight(10); // ~10 pt tall bar

  body.appendParagraph('');

  // Document title
  const titlePara = body.appendParagraph(data.documentTitle);
  titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titlePara.setSpacingAfter(4);
  titlePara.editAsText()
    .setFontFamily('Montserrat')
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);

  // Subtitle
  const subtitlePara = body.appendParagraph(data.subtitle);
  subtitlePara.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  subtitlePara.setSpacingAfter(24);
  subtitlePara.editAsText()
    .setFontFamily('Open Sans')
    .setFontSize(15)
    .setBold(false)
    .setForegroundColor(COLORS.SECONDARY);

  // Prepared for block
  appendCoverLabel(body, 'Prepared for:');
  appendCoverValue(body, data.clientName);
  appendCoverValue(body, data.clientCompany);

  body.appendParagraph('').setSpacingAfter(4);

  // Prepared by block
  appendCoverLabel(body, 'Prepared by:');
  appendCoverValue(body, data.yourName);
  appendCoverValue(body, data.yourCompany);

  body.appendParagraph('').setSpacingAfter(4);
  appendCoverLabel(body, `Date: ${data.date}`);

  // End cover page
  body.appendPageBreak();
}

function appendCoverLabel(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  p.setSpacingAfter(1);
  p.editAsText().setFontFamily('Open Sans').setFontSize(9).setForegroundColor(COLORS.TEXT_MUTED);
}

function appendCoverValue(body, text) {
  const p = body.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  p.setSpacingAfter(2);
  p.editAsText().setFontFamily('Open Sans').setFontSize(12).setForegroundColor(COLORS.TEXT_DARK).setBold(false);
}
```

---

## 6. Tables

### Table Design Principles

Tables are the hardest element to polish in Google Docs. The defaults are heavy and ugly — override everything intentionally.

### Header Row

- Background: `PRIMARY` color.
- Text: `#FFFFFF`, bold, 10 pt.
- Alignment: Left for text columns, Right for numeric columns.
- Padding: 6 pt top/bottom, 8 pt left/right minimum.
- Never let header text wrap — size columns wide enough for labels to fit on one line.

### Alternating Data Rows

- Odd rows: `#FFFFFF` (white).
- Even rows: `TABLE_ALT` (e.g., `#EEF3F9`).
- Text: `TEXT_DARK`, 10.5 pt.
- Padding: 5 pt top/bottom, 8 pt left/right.
- Skip alternating colors on tables with fewer than 4 data rows — it looks arbitrary.

### Border Style

- **Outer border:** 1 pt, `TABLE_BORDER` color.
- **Inner horizontal borders:** 0.5 pt, `TABLE_BORDER`.
- **Inner vertical borders:** None (0 pt) — removing verticals gives a clean, modern look.
- Never use the default thick black borders Docs inserts.

**Note:** `DocumentApp` can set `table.setBorderWidth(0)` to remove all borders, but fine-grained per-side cell border control requires the Docs REST API `updateTableCellStyle` request.

### Column Widths

Define column widths explicitly — never trust Docs auto-sizing. On a 6 in text column:

| Table Type | Description | Qty | Unit Price | Total |
|---|---|---|---|---|
| Invoice / Proposal | 50% (3 in) | 10% | 20% | 20% |
| Feature comparison | 40% | 20% | 20% | 20% |
| Simple 2-col | 60% | — | — | 40% |

### Totals Row

- Background: `ACCENT` color or a darkened `PRIMARY`.
- Text: `#FFFFFF`, bold, 12 pt.
- The totals row is the punchline of a financial table — make it visually unmissable.

### Apps Script: Table Styling

```javascript
// COLUMN_WIDTHS must sum to the body text width in points
// Body width = page width − left margin − right margin
// Letter with 1.25 in margins: 612 − 90 − 90 = 432 pt
const COLUMN_WIDTHS_INVOICE = [216, 43, 86, 87]; // 432 pt total

function buildStyledTable(body, headerRow, dataRows) {
  const allRows = [headerRow, ...dataRows];
  const table = body.appendTable(allRows);
  const numCols = headerRow.length;

  // Set column widths
  for (let row = 0; row < table.getNumRows(); row++) {
    for (let col = 0; col < numCols; col++) {
      table.getCell(row, col).setWidth(COLUMN_WIDTHS_INVOICE[col]);
    }
  }

  // Style header row
  const hRow = table.getRow(0);
  for (let col = 0; col < numCols; col++) {
    const cell = hRow.getCell(col);
    cell.setBackgroundColor(COLORS.TABLE_HEADER);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    cell.setPaddingTop(6);
    cell.setPaddingBottom(6);
    cell.setPaddingLeft(8);
    cell.setPaddingRight(8);
    cell.editAsText()
      .setFontFamily('Open Sans')
      .setFontSize(10)
      .setBold(true)
      .setForegroundColor(COLORS.WHITE);
  }

  // Style data rows with alternating colors
  for (let r = 1; r < table.getNumRows(); r++) {
    const bgColor = r % 2 === 0 ? COLORS.TABLE_ALT : COLORS.WHITE;
    for (let col = 0; col < numCols; col++) {
      const cell = table.getRow(r).getCell(col);
      cell.setBackgroundColor(bgColor);
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);
      cell.editAsText()
        .setFontFamily('Open Sans')
        .setFontSize(10.5)
        .setBold(false)
        .setForegroundColor(COLORS.TEXT_DARK);
    }
    // Right-align numeric columns (last two)
    table.getRow(r).getCell(numCols - 1)
      .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    table.getRow(r).getCell(numCols - 2)
      .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }

  // Remove all default borders
  table.setBorderWidth(0);

  return table;
}
```

---

## 7. Paragraph Styles

### Using Named Styles — Always

Google Docs has six heading levels plus Title, Subtitle, and Normal. Always assign the named heading style — do NOT manually style text to look like a heading without calling `setHeading()`. Named styles power the Table of Contents and the document outline panel.

| Style Name | `ParagraphHeading` Constant | Use |
|---|---|---|
| Normal | `NORMAL` | Body text |
| Title | `TITLE` | Document title (cover page only) |
| Subtitle | `SUBTITLE` | Document subtitle |
| Heading 1 | `HEADING1` | Major sections |
| Heading 2 | `HEADING2` | Subsections |
| Heading 3 | `HEADING3` | Sub-subsections |
| Heading 4 | `HEADING4` | Rarely needed |

### Paragraph Spacing Reference

**Never use blank paragraph breaks for vertical space.** Use `setSpacingBefore` / `setSpacingAfter`.

| Element | Space Before | Space After |
|---|---|---|
| H1 | 16 pt | 6 pt |
| H2 | 12 pt | 4 pt |
| H3 | 8 pt | 3 pt |
| Body paragraph | 0 pt | 7 pt |
| List item (non-final) | 0 pt | 2 pt |
| List item (final in list) | 0 pt | 8 pt |
| Table block | 6 pt before | 8 pt after |
| Image + caption block | 0 pt | 10 pt |

### Indentation

- **Body text:** No left indent (0).
- **Block quotes:** 36 pt (0.5 in) left indent, italic, `TEXT_MUTED` color.
- **Terms and conditions:** 18 pt (0.25 in) left indent, 9–10 pt, `TEXT_MUTED`.
- **List items:** Managed automatically by Docs list styles at 18 pt (level 1) and 36 pt (level 2).

### Apps Script: Paragraph Styling

```javascript
function appendH1(body, text) {
  const para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  para.setSpacingBefore(16);
  para.setSpacingAfter(6);
  // Always set font AFTER setHeading() — heading resets may override earlier font calls
  para.editAsText()
    .setFontFamily('Montserrat')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);
  return para;
}

function appendH2(body, text) {
  const para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  para.setSpacingBefore(12);
  para.setSpacingAfter(4);
  para.editAsText()
    .setFontFamily('Montserrat')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.SECONDARY);
  return para;
}

function appendH3(body, text) {
  const para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  para.setSpacingBefore(8);
  para.setSpacingAfter(3);
  para.editAsText()
    .setFontFamily('Montserrat')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.SECONDARY);
  return para;
}

function appendBodyPara(body, text) {
  const para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  para.setSpacingAfter(7);
  para.setLineSpacing(1.2);
  para.editAsText()
    .setFontFamily('Open Sans')
    .setFontSize(11)
    .setBold(false)
    .setForegroundColor(COLORS.TEXT_DARK);
  return para;
}
```

---

## 8. Lists

### Bullet vs. Numbered

| Use Bullets When | Use Numbers When |
|---|---|
| Items have no required order | Order or sequence matters |
| Features, benefits, inclusions | Steps in a process |
| Items are roughly equal in weight | Priority or ranking matters |

### Indent Levels

- **Level 1:** Standard bullet or number. Docs applies ~18 pt left indent automatically.
- **Level 2:** Sub-item. ~36 pt left indent. Use a different bullet style (dash or hollow circle).
- **Level 3+:** Avoid. Three-level nesting means the content hierarchy is a document structure problem, not a list problem. Promote the sub-items to their own H3 section instead.

### List Typography

- Same font and size as body text (11 pt, `Open Sans` or `Georgia`).
- `SpacingAfter` on non-final list items: **2–3 pt** (tighter than body paragraphs — lists read as a unit).
- `SpacingAfter` on the final list item: **8 pt** (separates the list block from what follows).

### Apps Script: Lists

```javascript
function appendBulletList(body, items) {
  items.forEach((text, index) => {
    const item = body.appendListItem(text);
    item.setGlyphType(DocumentApp.GlyphType.BULLET);
    item.setNestingLevel(0);
    item.setFontSize(11);
    item.setForegroundColor(COLORS.TEXT_DARK);
    item.setSpacingAfter(index === items.length - 1 ? 8 : 2);
  });
}

function appendNumberedList(body, items) {
  items.forEach((text, index) => {
    const item = body.appendListItem(text);
    item.setGlyphType(DocumentApp.GlyphType.DECIMAL);
    item.setNestingLevel(0);
    item.setFontSize(11);
    item.setForegroundColor(COLORS.TEXT_DARK);
    item.setSpacingAfter(index === items.length - 1 ? 8 : 3);
  });
}
```

---

## 9. Images

### Alignment

- **Full-width images** (diagrams, banners, screenshots): Center-aligned, inline, spanning the full text column.
- **Small icons / badges:** Inline with text, or right-aligned in a two-column borderless table next to descriptive text.
- **Logos on cover pages:** Left-aligned or center-aligned. Pick one and be consistent across all document types.

### Wrapping

Use `DocumentApp.WrapType.INLINE` for all script-generated images. Floating images (WRAP_TEXT, BREAK_TEXT) drift when documents are reopened or viewed on different devices. If you need an image next to text, use a two-column borderless table — the image goes in one cell, the text in the other.

### Sizing

- **Logo in header:** Max 0.5 in (36 pt) tall, proportional width.
- **Cover page logo:** Max 1.5 in (108 pt) tall, proportional width.
- **Body diagram / screenshot:** Max text column width (e.g., 432 pt for Letter with 1.25 in margins). Always scale proportionally.
- **Never stretch an image.** Set one dimension and calculate the other from the original aspect ratio.

### Captions

Add a caption as a separate Normal paragraph immediately below the image:

- Font: 9 pt, italic, `TEXT_MUTED` (`#6B7280`).
- Alignment: Center.
- Format: "Figure 1. Description of the image."
- `SpacingAfter`: 10 pt.

### Apps Script: Images

```javascript
function insertBodyImage(body, imageBlob, captionText, maxWidthPt = 432) {
  const image = body.appendImage(imageBlob);

  // Scale proportionally to max width
  const originalWidth  = image.getWidth();
  const originalHeight = image.getHeight();
  const scale = Math.min(1, maxWidthPt / originalWidth);
  image.setWidth(originalWidth * scale);
  image.setHeight(originalHeight * scale);

  // Center the image
  // Images are inline elements — align the containing paragraph
  // (In Apps Script, the image itself is a PositionedImage or inline child)

  // Caption
  const caption = body.appendParagraph(`Figure. ${captionText}`);
  caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  caption.setSpacingAfter(10);
  caption.editAsText()
    .setFontFamily('Open Sans')
    .setFontSize(9)
    .setItalic(true)
    .setForegroundColor(COLORS.TEXT_MUTED);
}
```

**Fetching images from Drive:**

```javascript
const LOGO_FILE_ID = 'YOUR_LOGO_FILE_ID_HERE';

function getLogoBlob() {
  return DriveApp.getFileById(LOGO_FILE_ID).getBlob();
}
```

---

## 10. Document Structure

### Standard Section Order

**Proposal:**
1. Cover Page
2. Table of Contents
3. Executive Summary
4. Scope of Work
5. Deliverables and Timeline
6. Pricing / Investment
7. Terms and Conditions
8. Signature Block

**Contract:**
1. Cover Page
2. Parties and Recitals
3. Definitions
4. Scope of Services
5. Compensation and Payment Terms
6. Term and Termination
7. Intellectual Property
8. Confidentiality
9. Limitation of Liability
10. General Provisions
11. Signature Block

**Invoice:**
1. Header (company info + client info)
2. Invoice metadata (number, date, due date)
3. Line items table
4. Subtotal / Tax / Total block
5. Payment instructions
6. Notes and Terms (brief)

### Table of Contents

Include a TOC only for documents longer than 4 pages. The TOC auto-generates from named heading styles — another strong reason to use `setHeading()` consistently.

```javascript
function appendTocPlaceholder(body) {
  // DocumentApp cannot insert a live TOC field programmatically
  // Insert a placeholder — user refreshes via Docs > Insert > Table of contents
  const heading = body.appendParagraph('Table of Contents');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const placeholder = body.appendParagraph(
    '[Update this table of contents: Insert > Table of contents in Google Docs]'
  );
  placeholder.editAsText().setItalic(true).setForegroundColor(COLORS.TEXT_MUTED);
  placeholder.setSpacingAfter(12);

  body.appendPageBreak();
}
// Advanced alternative: use Docs REST API insertTableOfContents batchUpdate request
```

### Signature Block

```
Agreed and accepted by both parties:

_________________________________      _________________________________
[Client Name]                          [Your Name]
[Client Title]                         [Your Title]
[Client Company]                       [Your Company]
Date: ___________________________      Date: ___________________________
```

Use a two-column, borderless table for side-by-side signature blocks. Signature line = a run of underscores at body font size.

```javascript
function appendSignatureBlock(body, clientData, yourData) {
  body.appendParagraph('').setSpacingBefore(24);
  body.appendParagraph('Agreed and accepted by both parties:')
    .editAsText().setFontSize(11).setForegroundColor(COLORS.TEXT_DARK);
  body.appendParagraph('').setSpacingAfter(12);

  const sigRows = [
    ['_________________________________', '_________________________________'],
    [clientData.name, yourData.name],
    [clientData.title, yourData.title],
    [clientData.company, yourData.company],
    ['Date: ___________________________', 'Date: ___________________________'],
  ];

  const sigTable = body.appendTable(sigRows);
  sigTable.setBorderWidth(0);

  for (let r = 0; r < sigTable.getNumRows(); r++) {
    for (let col = 0; col < 2; col++) {
      const cell = sigTable.getRow(r).getCell(col);
      const fontSize = r === 0 ? 11 : (r >= 3 ? 9 : 10);
      const color = r >= 1 && r <= 3 ? COLORS.TEXT_DARK : COLORS.TEXT_MUTED;
      cell.editAsText().setFontSize(fontSize).setForegroundColor(color);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(0).setPaddingRight(12);
    }
  }
}
```

---

## 11. Invoice / Proposal / Contract Layouts

### Invoice Layout

```
┌──────────────────────────────────────────────────────────┐
│  [LOGO]                          INVOICE                 │
│  Company Name                    Invoice #: INV-2026-047 │
│  Address Line 1                  Date: March 23, 2026    │
│  Phone | Email                   Due:  April 22, 2026    │
├──────────────────────────────────────────────────────────┤
│  BILL TO:                                                │
│  Client Name                                             │
│  Client Company                                          │
│  Client Address                                          │
├──────────────────────────────────────────────────────────┤
│  Description        Qty    Unit Price      Total         │
│  ──────────────────────────────────────────────────────  │
│  Website Design      1     $3,500.00    $3,500.00        │
│  SEO Package         1       $800.00      $800.00        │
│  Hosting Setup       1       $150.00      $150.00        │
├──────────────────────────────────────────────────────────┤
│                              Subtotal:   $4,450.00       │
│                              Tax (10%):    $445.00       │
│                           TOTAL DUE:    $4,895.00        │
├──────────────────────────────────────────────────────────┤
│  Payment Instructions                                    │
│  Bank / PayPal / Check details here                      │
├──────────────────────────────────────────────────────────┤
│  Terms: Payment due within 30 days of invoice date.      │
└──────────────────────────────────────────────────────────┘
```

### Totals Section

- Right-align all labels and values.
- **TOTAL DUE row:** `ACCENT` background, `WHITE` text, bold, 12 pt.
- Subtotal and tax rows: No background, 10.5 pt, `TEXT_DARK` values, `TEXT_MUTED` labels.
- Use a 2-column borderless table (label | value) — never try to right-align text within a single-cell paragraph.

```javascript
function appendTotalsSection(body, subtotal, taxRate) {
  const taxAmount = subtotal * taxRate;
  const total = subtotal + taxAmount;

  const rows = [
    ['Subtotal:', formatCurrency(subtotal)],
    [`Tax (${Math.round(taxRate * 100)}%):`, formatCurrency(taxAmount)],
    ['TOTAL DUE:', formatCurrency(total)],
  ];

  const table = body.appendTable(rows);
  table.setBorderWidth(0);

  // Subtotal and tax rows
  for (let r = 0; r < 2; r++) {
    table.getRow(r).getCell(0)
      .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .editAsText().setFontSize(10.5).setForegroundColor(COLORS.TEXT_MUTED);
    table.getRow(r).getCell(1)
      .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .editAsText().setFontSize(10.5).setForegroundColor(COLORS.TEXT_DARK);
  }

  // TOTAL DUE row
  const totalRow = table.getRow(2);
  for (let col = 0; col < 2; col++) {
    totalRow.getCell(col)
      .setBackgroundColor(COLORS.ACCENT)
      .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setPaddingTop(6).setPaddingBottom(6)
      .setPaddingLeft(10).setPaddingRight(10)
      .editAsText()
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(COLORS.WHITE);
  }

  return table;
}

function formatCurrency(amount) {
  return '$' + amount.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
```

### Proposal: Pricing Section

- Use the line items table (same style as invoice).
- After the table, add a standalone callout paragraph: "Investment Total: $X,XXX" in bold `ACCENT` or `PRIMARY` color at 13–14 pt.
- Follow pricing with a "What's included" bullet list.
- End the pricing section with a "Next Steps" paragraph or a callout box.

### Contract: Clauses and Terms

- Number each clause with a Heading 3 style: "3. Confidentiality" — not "3.0" or "Article III".
- Body text under each clause: 10.5–11 pt, 1.3 line spacing, `TEXT_DARK`.
- **Definitions section:** Use a 2-column borderless table. Left column = defined term (bold), right column = definition.
- **Confidentiality notice:** Insert at the very top of the document (before the cover page), centered, 9 pt, italic, `TEXT_MUTED`: "CONFIDENTIAL — This document contains proprietary information."
- **Important:** Use `body.insertParagraph(0, 'text')` to prepend the confidentiality notice before all other content.

---

## 12. What to Avoid

### Typography Mistakes

- **More than 2 typefaces.** Three fonts looks amateur. Four looks broken.
- **All-caps body text.** One level of all-caps (a single label or category) is acceptable. All-caps paragraphs are never acceptable.
- **Underlined text that is not a hyperlink.** Underline exclusively signals a link. Use bold for inline emphasis.
- **Justified alignment.** Google Docs justification produces uneven word spacing and awkward rivers of whitespace. Use left-aligned body text throughout.
- **Font sizes below 9 pt** for any content the client must read.

### Spacing Mistakes

- **Blank paragraph breaks for vertical space.** Use `setSpacingBefore` / `setSpacingAfter`. Blank paragraphs break headings from their content when pages reflow.
- **Inconsistent line spacing.** Pick one value (1.15 or 1.2) and apply it to all body text.
- **Tables with no cell padding.** Minimum padding: 5 pt top/bottom, 8 pt left/right. No padding makes tables feel cramped and hard to scan.
- **Content crowding page edges.** Never go below 0.5 in margins for any document type.

### Color Mistakes

- **Too many colors.** If you cannot name the role of each color, you have too many.
- **Low-contrast text.** Light gray on white (e.g., `#CCCCCC` on `#FFFFFF`) fails readability. Body text must be `#1A1A1A` or darker.
- **Colored body text.** Color belongs on headings, table headers, and callouts. Body copy must stay near-black.
- **Overused accent color.** If the accent appears on 8+ elements, it adds visual noise instead of emphasis. Accent = 1–2 specific, intentional uses per document.
- **Red for emphasis.** In financial documents, red means negative, error, or danger. Never use it for highlights.

### Structure Mistakes

- **Wall of text.** No paragraph should exceed 5–6 lines. Break long content into shorter paragraphs or a list.
- **Skipping heading levels.** Never jump from H1 to H3. The document outline must be hierarchically correct.
- **Orphaned totals or signature blocks.** Add explicit page breaks before these critical elements if there is any risk they will float to the top of a new page with no context.
- **Missing the cover page.** Client-facing proposals and contracts must have a cover page. Invoices do not need one.

---

## 13. Apps Script Implementation Notes

### Core API Objects

```javascript
// Create a new document
const doc = DocumentApp.create('Document Title');

// Or open an existing document
const doc = DocumentApp.openById(DOC_ID);

const body = doc.getBody();
const id   = doc.getId();
const url  = doc.getUrl();
```

### Appending Elements

```javascript
// Paragraph
const para = body.appendParagraph('Text here');

// Heading (always set heading BEFORE font styling)
const h1 = body.appendParagraph('Section Title');
h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
// Then apply font overrides here...

// Page break
body.appendPageBreak();

// Horizontal rule
body.appendHorizontalRule();

// Table from 2D string array
const table = body.appendTable([
  ['Header 1', 'Header 2', 'Header 3'],
  ['Row 1 A',  'Row 1 B',  'Row 1 C'],
]);

// List item
const item = body.appendListItem('Bullet point text');
item.setGlyphType(DocumentApp.GlyphType.BULLET);
item.setNestingLevel(0); // 0 = first level
```

### Text Styling

```javascript
// editAsText() returns a Text object for the entire paragraph
const text = para.editAsText();
text.setFontFamily('Open Sans');
text.setFontSize(11);
text.setForegroundColor('#1A1A1A');
text.setBold(false);
text.setItalic(false);
text.setUnderline(false);

// Style a specific character range (start index, end index, value)
text.setBold(0, 4, true);                    // bold first 5 chars
text.setForegroundColor(5, 10, '#2E6DA4');   // color chars 5–10

// appendText() adds text and returns the Text object
const t = para.appendText(' additional text');
```

### Paragraph Attributes

```javascript
para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

para.setSpacingBefore(12); // in points
para.setSpacingAfter(6);
para.setLineSpacing(1.2);  // multiplier
para.setIndentStart(18);   // left indent in points (18 pt ≈ 0.25 in)
para.setIndentEnd(18);     // right indent
```

### Table Cell Attributes

```javascript
const cell = table.getCell(rowIndex, colIndex);

cell.setBackgroundColor('#1B3A6B');
cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
cell.setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT);
cell.setPaddingTop(6);
cell.setPaddingBottom(6);
cell.setPaddingLeft(8);
cell.setPaddingRight(8);
cell.setWidth(216); // points

// Style text within the cell
cell.editAsText()
  .setFontFamily('Open Sans')
  .setFontSize(10)
  .setBold(true)
  .setForegroundColor('#FFFFFF');
```

### Critical: setHeading() Resets Font Styling

When you call `setHeading()`, Google Docs applies the default named style and may reset font family, size, and color. Always apply font overrides **after** the `setHeading()` call:

```javascript
// WRONG — font may be overwritten by setHeading()
para.editAsText().setFontFamily('Montserrat').setFontSize(18);
para.setHeading(DocumentApp.ParagraphHeading.HEADING1);

// CORRECT — heading first, then font overrides
para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
para.editAsText().setFontFamily('Montserrat').setFontSize(18).setForegroundColor(COLORS.PRIMARY);
```

### Inserting at a Specific Position

```javascript
// appendParagraph() always adds to the end of the body
// insertParagraph() inserts at a specific child index
const para = body.insertParagraph(0, 'This appears at the top');

// To insert after a specific element, get its index first
const index = body.getChildIndex(existingElement);
const newPara = body.insertParagraph(index + 1, 'Inserted after existingElement');
```

### Sharing and Saving

```javascript
// Always call saveAndClose() at the end — some changes may not persist without it
doc.saveAndClose();

// Get the shareable URL after saving
const url = `https://docs.google.com/document/d/${doc.getId()}/edit`;

// Move to a specific Drive folder
DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(FOLDER_ID));
```

### Document Metadata

```javascript
// Rename the document
doc.setName('Proposal — Acme Corp — March 2026');

// Add description (requires DriveApp)
DriveApp.getFileById(doc.getId()).setDescription('Auto-generated by TAK Scripts');
```

### Common Pitfalls

| Pitfall | Fix |
|---|---|
| `setHeading()` resets fonts | Apply font overrides after `setHeading()`, not before |
| Blank paragraphs for spacing | Use `setSpacingAfter()` / `setSpacingBefore()` |
| Images drift on reopening | Use `INLINE` wrapping only; use tables for image+text layouts |
| Table borders look wrong | Use `table.setBorderWidth(0)` then Docs REST API for per-side borders |
| `appendParagraph` at wrong position | Use `insertParagraph(index, text)` with explicit index |
| Script times out on large docs | Batch writes; avoid redundant `doc.getBody()` calls; split into functions |
| Changes don't persist | Always call `doc.saveAndClose()` before sharing the URL |
| Page numbers are static | Use Docs REST API `insertPageNumber` batchUpdate request for live fields |

### Full Document Generation Pattern

```javascript
function generateDocument(data) {
  const doc = DocumentApp.create(data.title);
  const body = doc.getBody();
  body.clear();

  setPageSetup(doc);

  // Cover page
  buildCoverPage(body, data);

  // TOC placeholder
  appendTocPlaceholder(body);

  // Body sections
  appendH1(body, 'Executive Summary');
  appendBodyPara(body, data.summary);

  appendH1(body, 'Scope of Work');
  appendBodyPara(body, data.scope);
  appendBulletList(body, data.deliverables);

  appendH1(body, 'Pricing');
  buildStyledTable(body, data.headers, data.lineItems);
  appendTotalsSection(body, data.subtotal, data.taxRate);

  appendH1(body, 'Terms and Conditions');
  appendBodyPara(body, data.terms);

  // Signature block
  appendSignatureBlock(body, data.client, data.company);

  // Header and footer
  addHeaderFooter(doc, data.title);

  doc.saveAndClose();
  return `https://docs.google.com/document/d/${doc.getId()}/edit`;
}
```

---

*Reference version: March 2026. Applies to Google Workspace (Docs, Apps Script V8 runtime). All hex colors are sRGB. Font availability subject to Google Fonts library. DocumentApp API limitations noted throughout — use Docs REST API via Advanced Google Services for features not available in DocumentApp.*
