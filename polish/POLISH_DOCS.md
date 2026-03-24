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
| Reports (dense data) | 0.75 in | 0.75 in | 0.75 in | 0.75 in | Narrow, use only when data density demands it |
| Cover Pages | 1.5 in | 1.5 in | 1.5 in | 1.5 in | Generous whitespace reads as premium |

**Rule of thumb:** Default to 1 in margins. Only go narrow if you have a specific space problem. Wider margins on cover pages make the document feel intentional, not squeezed.

### Page Size

Always use **Letter (8.5 × 11 in)** for US clients. Use **A4 (8.27 × 11.69 in)** only if the client is outside North America. Mixing sizes in a multi-section document is never appropriate.

### Orientation

- **Portrait** — default for all document types.
- **Landscape** — acceptable only for wide financial tables or architectural drawings embedded as a section. Never use landscape for an entire proposal or contract.

### Apps Script: Page Setup

```javascript
function setPageSetup(doc) {
  const body = doc.getBody();

  // Standard margins for proposals/contracts (in points: 1 in = 72 pt)
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(90);    // 1.25 in
  body.setMarginRight(90);   // 1.25 in

  // Narrow margins for invoices
  // body.setMarginTop(54);   // 0.75 in
  // body.setMarginLeft(72);
  // body.setMarginRight(72);
}
```

---

## 2. Typography

### Recommended Font Pairings

Google Docs is limited to fonts available in the Google Fonts library. These pairings produce professional, readable documents:

#### Pairing A — Clean & Modern (default recommendation)
| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Montserrat** | Bold (700) | See sizes below |
| Body | **Open Sans** | Regular (400) | 11 pt |
| Body emphasis | **Open Sans** | SemiBold (600) | 11 pt |
| Monospace / code | **Roboto Mono** | Regular | 10 pt |

#### Pairing B — Warm & Professional (good for law firms, consultants)
| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Roboto** | Bold (700) | See sizes below |
| Body | **Georgia** | Regular | 11 pt |
| Body emphasis | **Georgia** | Bold | 11 pt |

#### Pairing C — Minimal / Tech
| Role | Font | Weight | Size |
|---|---|---|---|
| Headings | **Lato** | Bold | See sizes below |
| Body | **Lato** | Regular | 11 pt |

**Never mix more than two typefaces in a single document.** Headings + body = 2 fonts maximum.

### Heading Size Scale

| Level | Style Name in Docs | Font Size | Weight | Case |
|---|---|---|---|---|
| Document Title | Title | 24–28 pt | Bold | Title Case |
| Subtitle | Subtitle | 14–16 pt | Regular or Light | Title Case |
| H1 (major sections) | Heading 1 | 18 pt | Bold | Title Case |
| H2 (subsections) | Heading 2 | 14 pt | Bold | Title Case |
| H3 (sub-subsections) | Heading 3 | 12 pt | Bold or SemiBold | Sentence case |
| H4 (rarely needed) | Heading 4 | 11 pt | Bold | Sentence case |

**Scale down consistently.** There should be a visible size difference between each level. If H1 and H2 look the same size, one of them is wrong.

### Body Text

- **Size:** 11 pt for most documents. Use 10.5 pt only if fitting to a page is critical (e.g., a one-page invoice).
- **Line spacing:** 1.15 is the Google Docs default and is acceptable. For proposals, bump to **1.2–1.3** for readability.
- **Paragraph spacing:** Add **6–8 pt after** each paragraph. This replaces blank lines between paragraphs — never use empty paragraph breaks for spacing.
- **Color:** `#222222` or `#1A1A1A` for body. Pure black (`#000000`) is technically fine but can feel harsh on screen.

### Line Length

At 8.5 in wide with 1.25 in margins, your text column is **6 in** — approximately 65–75 characters at 11 pt. This is ideal. Narrower columns increase hyphenation issues; wider columns fatigue the eye.

---

## 3. Color System

### Brand Color Roles

Define a small, intentional palette before writing any script. Every color in the document should serve a role.

```
PRIMARY      — Used for H1 headings, cover page elements, key accents
SECONDARY    — Used for H2/H3 headings, subheadings
ACCENT       — Used for call-to-action text, highlights, dividers
TABLE_HEADER — Used for table header row backgrounds
TABLE_ALT    — Alternating row tint (subtle, never distracting)
TEXT_DARK    — Primary body text
TEXT_MUTED   — Captions, footnotes, secondary labels
BORDER       — Table borders, dividers
WHITE        — Text on dark backgrounds
```

### Example Palette — Professional Blue

```javascript
const COLORS = {
  PRIMARY:      '#1B3A6B',  // Deep navy — H1, cover page, table headers
  SECONDARY:    '#2E6DA4',  // Medium blue — H2, H3, accents
  ACCENT:       '#E87722',  // Warm orange — highlights, totals row, CTA
  TABLE_HEADER: '#1B3A6B',  // Same as primary
  TABLE_ALT:    '#EEF3F9',  // Very light blue tint — alternating rows
  TABLE_BORDER: '#C5D5E8',  // Soft blue-gray border
  TEXT_DARK:    '#1A1A1A',  // Near-black body text
  TEXT_MUTED:   '#6B7280',  // Gray — captions, dates, secondary info
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
  TABLE_ALT:    '#EDF7F1',
  TABLE_BORDER: '#B7D9C6',
  TEXT_DARK:    '#1A1A1A',
  TEXT_MUTED:   '#6B7280',
  DIVIDER:      '#D1D5DB',
  WHITE:        '#FFFFFF',
};
```

### Color Rules

- **Text on a dark background** must always be `#FFFFFF` or near-white. Never use `#1A1A1A` on `#1B3A6B`.
- **Alternating row tint** should have a contrast ratio below 1.2:1 against white — it must be a whisper, not a stripe.
- **Accent color** should appear sparingly — totals, callout boxes, a single horizontal divider. If it appears everywhere, it stops being an accent.
- **Never use red** for emphasis in financial documents (it implies negative values). Use the accent color or bold text instead.

---

## 4. Headers and Footers

### Header Content

| Document Type | Header Left | Header Center | Header Right |
|---|---|---|---|
| Proposal | Company logo | — | Document title |
| Contract | Company name (text) | — | "CONFIDENTIAL" label |
| Invoice | Company logo | — | Invoice # and date |
| Report | Report title | — | Company name |

**Logo placement:** Left-aligned is standard. Center-aligned logos work on cover pages but look odd in repeating headers — keep the repeating header logo small (0.4–0.6 in tall, proportional width).

**Page numbers:** Always right-aligned in the footer. Format: `Page X of Y` for multi-page documents. Invoice page numbers are optional if single-page.

### Footer Content

| Left | Center | Right |
|---|---|---|
| Company address or tagline | — | Page X of Y |
| Confidentiality notice (contracts) | — | Page X of Y |
| Document date | — | Page X of Y |

**Font size in header/footer:** 8–9 pt. Use `TEXT_MUTED` color (`#6B7280`). The header/footer should recede — it's structural, not readable content.

### Apps Script: Headers and Footers

```javascript
function addHeaderFooter(doc, docTitle, invoiceNumber) {
  // --- HEADER ---
  const header = doc.addHeader();
  const headerTable = header.appendTable([[' ', docTitle]]);

  // Style left cell (logo placeholder or company name)
  const leftCell = headerTable.getCell(0, 0);
  leftCell.setText('Your Company Name');
  leftCell.setFontSize(9).setForegroundColor('#6B7280');

  // Style right cell (document title)
  const rightCell = headerTable.getCell(0, 1);
  rightCell.setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rightCell.setFontSize(9).setForegroundColor('#6B7280');

  // Remove table borders from header table
  headerTable.setBorderWidth(0);

  // --- FOOTER ---
  const footer = doc.addFooter();
  const footerPara = footer.appendParagraph('');
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  footerPara.appendText('Page ');
  // Note: dynamic page numbers require a page number field
  footerPara.setFontSize(8);
  footerPara.setForegroundColor('#6B7280');
}
```

**Note on dynamic page numbers:** `DocumentApp` does not currently support inserting live `PAGE` / `NUMPAGES` fields via script. For documents where page count matters, you can insert a static placeholder and note it in your documentation, or use the Docs REST API's `insertInlineImage` / batch update endpoint which does support page number fields.

---

## 5. Cover Pages

### Professional Cover Page Layout

A cover page should communicate: who you are, what this document is, and create a strong first impression. Use generous whitespace.

```
┌─────────────────────────────────┐
│                                 │
│   [Logo — top left or center]   │
│                                 │
│                                 │
│   [Decorative bar or divider]   │
│                                 │
│   DOCUMENT TITLE                │
│   Document Subtitle             │
│                                 │
│   Prepared for:                 │
│   Client Name                   │
│   Client Company                │
│                                 │
│   Prepared by:                  │
│   Your Name                     │
│   Your Company                  │
│                                 │
│   Date: March 23, 2026          │
│                                 │
└─────────────────────────────────┘
```

### Cover Page Design Rules

- **One accent element** — a colored horizontal bar (full-width, ~0.15 in thick) in `PRIMARY` color is clean and striking.
- **Title at 28–32 pt**, bold, in `PRIMARY` color.
- **Subtitle at 16 pt**, regular, in `SECONDARY` or `TEXT_MUTED`.
- **"Prepared for/by" labels** at 9 pt, `TEXT_MUTED`. Client/company names at 12 pt, `TEXT_DARK`.
- **Date** at 10 pt, `TEXT_MUTED`.
- **No page number** on the cover page (suppress with a section break or by starting page numbering from page 2).
- Always end the cover page with an explicit page break before the table of contents or first content section.

### Apps Script: Cover Page

```javascript
function buildCoverPage(body, data) {
  // Company name as logo placeholder
  const companyPara = body.appendParagraph(data.companyName);
  companyPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  companyPara.getChild(0).asText()
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);

  // Spacer paragraphs — use sparingly, only on cover pages
  body.appendParagraph('');
  body.appendParagraph('');

  // Colored accent bar (simulated with a styled table row)
  const accentTable = body.appendTable([['']]);
  accentTable.getRow(0).getCell(0)
    .setBackgroundColor(COLORS.PRIMARY);
  accentTable.setBorderWidth(0);
  // Set row height via minimum height (approx 10pt = narrow bar)
  accentTable.getRow(0).setMinimumHeight(10);

  body.appendParagraph('');

  // Document title
  const titlePara = body.appendParagraph(data.documentTitle);
  titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titlePara.getChild(0).asText()
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  // Subtitle
  const subtitlePara = body.appendParagraph(data.subtitle);
  subtitlePara.getChild(0).asText()
    .setFontSize(15)
    .setBold(false)
    .setForegroundColor(COLORS.SECONDARY);

  body.appendParagraph('');
  body.appendParagraph('');

  // Prepared for block
  appendLabel(body, 'Prepared for:');
  appendValue(body, data.clientName);
  appendValue(body, data.clientCompany);

  body.appendParagraph('');

  // Prepared by block
  appendLabel(body, 'Prepared by:');
  appendValue(body, data.yourName);
  appendValue(body, data.yourCompany);

  body.appendParagraph('');
  appendLabel(body, `Date: ${data.date}`);

  // Page break to end cover page
  body.appendPageBreak();
}

function appendLabel(body, text) {
  const p = body.appendParagraph(text);
  p.getChild(0).asText().setFontSize(9).setForegroundColor(COLORS.TEXT_MUTED);
}

function appendValue(body, text) {
  const p = body.appendParagraph(text);
  p.getChild(0).asText().setFontSize(12).setForegroundColor(COLORS.TEXT_DARK).setBold(false);
}
```

---

## 6. Tables

### Table Design Principles

Tables are the hardest element to get right in Google Docs. The defaults are ugly — you must override everything intentionally.

### Header Row

- Background: `PRIMARY` color (`#1B3A6B`).
- Text: `#FFFFFF`, bold, 10 pt.
- Alignment: Center or Left (Left preferred for text columns, Center for numeric columns).
- Never let the header row wrap — keep column widths wide enough to show labels on one line.

### Alternating Rows

- Odd rows: `#FFFFFF` (white)
- Even rows: `TABLE_ALT` (`#EEF3F9`)
- Text: `TEXT_DARK` (`#1A1A1A`), 10.5–11 pt
- Never use alternating rows on a table with fewer than 4 data rows — it looks odd.

### Borders

- **Outer border:** 1 pt, `TABLE_BORDER` color (`#C5D5E8`).
- **Inner horizontal borders:** 0.5 pt, `TABLE_BORDER`.
- **Inner vertical borders:** None (0 pt) — removing vertical borders gives a clean, modern look.
- **Never** use the default thick black borders that Google Docs inserts.

### Column Widths

- Define column widths explicitly — never let Docs auto-size.
- For a 6 in text column with 4 columns: decide the split based on content type. A description column should be wider than a quantity column.
- **Invoice line item example:** Description 50%, Qty 10%, Unit Price 20%, Total 20%.

### Totals Row

- Background: `ACCENT` color or a slightly darker shade of `PRIMARY`.
- Text: `#FFFFFF`, bold.
- Make the totals row visually distinct — it is the punchline of a financial table.

### Apps Script: Table Styling

```javascript
function styleTable(table, headers, rows) {
  const COLUMN_WIDTHS = [230, 50, 100, 100]; // points, must sum to body width
  const NUM_COLS = headers.length;

  // Set column widths
  for (let col = 0; col < NUM_COLS; col++) {
    for (let row = 0; row < table.getNumRows(); row++) {
      table.getCell(row, col).setWidth(COLUMN_WIDTHS[col]);
    }
  }

  // Style header row
  const headerRow = table.getRow(0);
  for (let col = 0; col < NUM_COLS; col++) {
    const cell = headerRow.getCell(col);
    cell.setBackgroundColor(COLORS.TABLE_HEADER);
    cell.setBold(true);
    cell.setForegroundColor(COLORS.WHITE);
    cell.setFontSize(10);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    cell.setPaddingTop(6);
    cell.setPaddingBottom(6);
    cell.setPaddingLeft(8);
    cell.setPaddingRight(8);
  }

  // Style data rows with alternating colors
  for (let r = 1; r < table.getNumRows(); r++) {
    const isAlt = r % 2 === 0;
    const bgColor = isAlt ? COLORS.TABLE_ALT : COLORS.WHITE;
    for (let col = 0; col < NUM_COLS; col++) {
      const cell = table.getRow(r).getCell(col);
      cell.setBackgroundColor(bgColor);
      cell.setForegroundColor(COLORS.TEXT_DARK);
      cell.setFontSize(10.5);
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);
    }
  }

  // Remove borders (set all to 0 pt, then add outer border)
  table.setBorderWidth(0);
  // Google Docs API note: individual cell border control is limited in DocumentApp.
  // For fine border control, use the Docs REST API batchUpdate with tableCellBorder updates.
}
```

---

## 7. Paragraph Styles

### Using Named Styles

Google Docs has six heading levels plus Title, Subtitle, and Normal. Always use these named styles — do NOT manually style text to look like a heading without assigning the heading style. Named styles are what power the Table of Contents and document outline.

| Style Name | `ParagraphHeading` Constant | Typical Use |
|---|---|---|
| Normal | `NORMAL` | Body text |
| Title | `TITLE` | Document title (cover page only) |
| Subtitle | `SUBTITLE` | Document subtitle |
| Heading 1 | `HEADING1` | Major section (1. Scope of Work) |
| Heading 2 | `HEADING2` | Subsection (1.1 Deliverables) |
| Heading 3 | `HEADING3` | Sub-subsection |
| Heading 4 | `HEADING4` | Rarely used; avoid if possible |

### Paragraph Spacing

**Do not use blank paragraph breaks to create vertical space.** Instead, set `SpacingAfter` on paragraphs.

| Element | Space Before | Space After |
|---|---|---|
| H1 | 16 pt | 6 pt |
| H2 | 12 pt | 4 pt |
| H3 | 8 pt | 3 pt |
| Body paragraph | 0 pt | 6–8 pt |
| List item | 0 pt | 2–3 pt |
| Table (block) | 8 pt | 8 pt |

### Indentation

- **Body text:** No left indent (0 in).
- **Block quotes:** 0.5 in left indent, `TEXT_MUTED` color, italic.
- **Terms & conditions text:** 0.25 in left indent, 9 pt, `TEXT_MUTED`.
- **List items:** First level 0.25 in, second level 0.5 in (handled automatically by Docs list styles).

### Apps Script: Paragraph Spacing

```javascript
function styleHeading1(para) {
  para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  para.setSpacingBefore(16);
  para.setSpacingAfter(6);
  para.getChild(0).asText()
    .setFontFamily('Montserrat')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(COLORS.PRIMARY);
}

function styleHeading2(para) {
  para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  para.setSpacingBefore(12);
  para.setSpacingAfter(4);
  para.getChild(0).asText()
    .setFontFamily('Montserrat')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.SECONDARY);
}

function styleBodyParagraph(para) {
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  para.setSpacingAfter(7);
  para.setLineSpacing(1.2);
  para.getChild(0).asText()
    .setFontFamily('Open Sans')
    .setFontSize(11)
    .setBold(false)
    .setForegroundColor(COLORS.TEXT_DARK);
}
```

---

## 8. Lists

### Bullet vs. Numbered

| Use Bullets When... | Use Numbers When... |
|---|---|
| Items have no required order | Order or sequence matters |
| Features, benefits, inclusions | Steps in a process |
| Items are roughly equal in weight | Priority ranking matters |
| 3–8 items | Any count (numbers handle long lists better) |

### Indent Levels

- **Level 1:** Standard bullet or number. Left indent 0.25 in.
- **Level 2:** Sub-item. Left indent 0.5 in. Use a different bullet style (dash or hollow circle).
- **Level 3+:** Avoid. If you are nesting 3 levels deep, your content hierarchy is a document structure problem, not a list problem.

### List Typography

- Same font and size as body text (11 pt Open Sans or Georgia).
- `SpacingAfter` on list items: **2–3 pt** (tighter than body paragraphs — lists are scannable and should read as a unit).
- Add `SpacingAfter` of **8 pt** on the final list item to separate the list from what follows.

### Apps Script: Lists

```javascript
function appendBulletList(body, items) {
  items.forEach((item, index) => {
    const listItem = body.appendListItem(item);
    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    listItem.setNestingLevel(0);
    listItem.setFontSize(11);
    listItem.setForegroundColor(COLORS.TEXT_DARK);
    // Tighter spacing within list, more space after the last item
    listItem.setSpacingAfter(index === items.length - 1 ? 8 : 2);
  });
}

function appendNumberedList(body, items) {
  items.forEach((item, index) => {
    const listItem = body.appendListItem(item);
    listItem.setGlyphType(DocumentApp.GlyphType.DECIMAL);
    listItem.setNestingLevel(0);
    listItem.setFontSize(11);
    listItem.setForegroundColor(COLORS.TEXT_DARK);
    listItem.setSpacingAfter(index === items.length - 1 ? 8 : 3);
  });
}
```

---

## 9. Images

### Alignment

- **Full-width images** (diagrams, banners): Center-aligned, inline, spanning the full text column.
- **Small icons / badges**: Inline with text, or right-aligned with text wrapping.
- **Logos on cover pages**: Left-aligned or center-aligned (pick one and be consistent).

### Wrapping

`DocumentApp.WrapType.INLINE` is the only wrapping mode that reliably works in Apps Script. Avoid floating images generated via script — they drift on re-open and across devices.

For images that must float next to text, use a two-column table (no borders) with text in one cell and the image in the other.

### Sizing

- **Logo in header:** Max 0.5 in tall, proportional width.
- **Cover page logo:** Max 1.5 in tall, proportional width.
- **Diagram / screenshot in body:** Max text column width (6 in). Always constrain proportionally.
- **Never stretch an image** (set only one dimension and let the other scale).

### Captions

Add a caption as a separate Normal paragraph immediately below the image:

- Font: 9 pt, italic, `TEXT_MUTED` (`#6B7280`).
- Alignment: Center.
- Text: "Figure 1. Description of the image."
- `SpacingAfter`: 10 pt (to visually separate the image+caption block from following content).

### Apps Script: Images

```javascript
function insertImageWithCaption(body, imageBlob, captionText, widthPt) {
  const image = body.appendImage(imageBlob);
  // Scale proportionally to desired width
  const originalWidth = image.getWidth();
  const originalHeight = image.getHeight();
  const scale = widthPt / originalWidth;
  image.setWidth(widthPt);
  image.setHeight(originalHeight * scale);

  // Caption paragraph
  const caption = body.appendParagraph(`Figure. ${captionText}`);
  caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  caption.getChild(0).asText()
    .setFontSize(9)
    .setItalic(true)
    .setForegroundColor(COLORS.TEXT_MUTED);
  caption.setSpacingAfter(10);
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
5. Deliverables & Timeline
6. Pricing / Investment
7. Terms & Conditions
8. Signature Block

**Contract:**
1. Cover Page
2. Parties and Recitals
3. Definitions
4. Scope of Services
5. Compensation & Payment Terms
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
4. Subtotal / Tax / Total
5. Payment instructions
6. Notes / Terms (brief)

### Table of Contents

Insert a TOC only for documents longer than 4 pages. In Google Docs, the TOC is auto-generated from Heading styles — another reason to use named styles consistently.

```javascript
// Insert a TOC — Docs will populate it from named headings
body.appendParagraph('Table of Contents')
  .setHeading(DocumentApp.ParagraphHeading.HEADING1);
body.appendParagraph('[Table of contents will appear here — use Docs Insert > Table of Contents to generate]')
  .getChild(0).asText().setItalic(true).setForegroundColor(COLORS.TEXT_MUTED);
body.appendPageBreak();
```

**Note:** `DocumentApp` cannot programmatically insert a live TOC field. Insert a placeholder and instruct the user to refresh it, or use the Docs REST API `insertTableOfContents` operation.

### Signature Block

```
Agreed and accepted:

__________________________________        __________________________________
[Client Name]                             [Your Name]
[Client Title]                            [Your Title]
[Client Company]                          [Your Company]
Date: ____________________________        Date: ____________________________
```

- Use a two-column, borderless table for side-by-side signatures.
- Signature line: a paragraph of underscores (`_______________`) at 11 pt.
- Label text (name, title, company): 10 pt, `TEXT_DARK`.
- "Date:" field: 10 pt, `TEXT_MUTED`.
- Add `SpacingBefore` of 24 pt on the signature block to give visual breathing room.

---

## 11. Invoice / Proposal / Contract Layouts

### Invoice Layout

```
┌──────────────────────────────────────────────────────┐
│  COMPANY LOGO             INVOICE                    │
│  Company Name             Invoice #: INV-2026-047    │
│  Address                  Date: March 23, 2026       │
│  Phone | Email            Due: April 22, 2026        │
├──────────────────────────────────────────────────────┤
│  BILL TO:                                            │
│  Client Name                                         │
│  Client Company                                      │
│  Client Address                                      │
├──────────────────────────────────────────────────────┤
│  [LINE ITEMS TABLE]                                  │
│  Description       Qty   Unit Price    Total         │
│  ─────────────────────────────────────────────────   │
│  Website Design     1    $3,500.00    $3,500.00      │
│  SEO Package        1    $800.00      $800.00        │
│  Hosting Setup      1    $150.00      $150.00        │
├──────────────────────────────────────────────────────┤
│                          Subtotal:    $4,450.00      │
│                          Tax (10%):   $445.00        │
│                          TOTAL DUE:  $4,895.00       │
├──────────────────────────────────────────────────────┤
│  Payment Instructions                                │
│  Bank Transfer / PayPal / Check details              │
├──────────────────────────────────────────────────────┤
│  Terms: Payment due within 30 days.                  │
└──────────────────────────────────────────────────────┘
```

### Invoice Totals Section

The totals section is the most important part of the invoice. Style rules:
- Right-align all labels and values.
- `TOTAL DUE` row: `ACCENT` background, `WHITE` text, bold, 12 pt.
- Subtotal and tax rows: No background, 10.5 pt, `TEXT_DARK`.
- Use a 2-column table (label | value) for the totals block, no visible borders.

```javascript
function appendTotalsSection(body, subtotal, taxRate, total) {
  const taxAmount = subtotal * taxRate;
  const rows = [
    ['Subtotal:', formatCurrency(subtotal)],
    [`Tax (${(taxRate * 100).toFixed(0)}%):`, formatCurrency(taxAmount)],
    ['TOTAL DUE:', formatCurrency(total)],
  ];

  const table = body.appendTable(rows);
  table.setBorderWidth(0);

  // Style subtotal and tax rows
  for (let r = 0; r < 2; r++) {
    table.getRow(r).getCell(0).setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setFontSize(10.5).setForegroundColor(COLORS.TEXT_MUTED);
    table.getRow(r).getCell(1).setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setFontSize(10.5).setForegroundColor(COLORS.TEXT_DARK);
  }

  // Style TOTAL row
  const totalRow = table.getRow(2);
  totalRow.getCell(0)
    .setBackgroundColor(COLORS.ACCENT)
    .setForegroundColor(COLORS.WHITE)
    .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setFontSize(12).setBold(true).setPaddingLeft(8).setPaddingRight(8);
  totalRow.getCell(1)
    .setBackgroundColor(COLORS.ACCENT)
    .setForegroundColor(COLORS.WHITE)
    .setHorizontalAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setFontSize(12).setBold(true).setPaddingLeft(8).setPaddingRight(8);
}

function formatCurrency(amount) {
  return `$${amount.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',')}`;
}
```

### Proposal: Pricing Section

- Use a table for itemized pricing (same style as invoice line items).
- After the table, add a callout paragraph: "Investment Total: $X,XXX" in a box or with a strong visual treatment.
- Follow pricing with a clear "What's included" bullet list.
- Add a "Next steps" paragraph or callout at the end of the pricing section.

### Contract: Terms Section

- Use numbered paragraphs (Heading 3 + body) for each clause.
- Body text at 10.5 pt, `TEXT_DARK`, line spacing 1.3.
- Definitions: use a two-column borderless table (term | definition).
- Confidentiality notice at the top of the document, 9 pt, centered, `TEXT_MUTED`, italic.

---

## 12. What to Avoid

### Typography Mistakes

- **More than 2 fonts.** Three fonts looks amateur. Four fonts looks broken.
- **All caps headings throughout the document.** One level of all-caps (e.g., H1 only) is acceptable. All-caps body text is never acceptable.
- **Underlined text that is not a hyperlink.** Underline signals a link. Use bold for emphasis.
- **Justified text.** Google Docs justification produces ugly, uneven word spacing. Use left-aligned body text.
- **Font sizes below 9 pt** for anything the client must read.

### Spacing Mistakes

- **Blank paragraph breaks as spacing.** Use `setSpacingBefore` / `setSpacingAfter`. Blank paragraphs break headings from their content when pages reflow.
- **Inconsistent line spacing.** Pick one value (1.15 or 1.2) and apply it everywhere.
- **Tables with no internal padding.** Cell padding of at least 5–6 pt top/bottom, 8 pt left/right is required for readability.
- **Content too close to page edges.** Never go below 0.5 in margins, even for the narrowest invoices.

### Color Mistakes

- **Too many colors.** If you cannot explain what each color is for, you have too many.
- **Low contrast text.** Light gray text on white background (`#CCCCCC` on `#FFFFFF`) fails readability. Keep body text at `#1A1A1A` minimum.
- **Colored body text.** Reserve color for headings, table headers, and callouts. Body text must always be near-black.
- **Random accent colors.** Your accent should be used for 1–2 specific purposes only (e.g., totals row, CTA). If it appears on 10 different elements, it adds noise, not emphasis.

### Structure Mistakes

- **Wall of text.** No paragraph should exceed 5–6 lines. Break long content into shorter paragraphs or use a list.
- **Skipping heading levels.** Never jump from H1 to H3. The outline must be hierarchically correct.
- **Tables for layout.** Using invisible tables to position text is a workaround, not a design — use it only when the alternative is worse.
- **Missing section breaks.** Without explicit page breaks, long documents reflow unpredictably and critical elements (signature blocks, totals) can orphan at the top of a page.

---

## 13. Apps Script Implementation Notes

### Core API Objects

```javascript
const doc = DocumentApp.create('Document Title');
// or open existing:
const doc = DocumentApp.openById(DOC_ID);

const body = doc.getBody();
```

### Appending Elements

```javascript
// Paragraph
const para = body.appendParagraph('Text here');

// Heading
const h1 = body.appendParagraph('Section Title');
h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);

// Page break
body.appendPageBreak();

// Horizontal rule
body.appendHorizontalRule();

// Table from 2D array
const data = [['Header 1', 'Header 2'], ['Row 1', 'Data']];
const table = body.appendTable(data);

// List item
const listItem = body.appendListItem('Item text');
listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
```

### Text Styling

```javascript
// Style a range within a paragraph
const text = para.editAsText();
text.setFontFamily('Open Sans');
text.setFontSize(11);
text.setForegroundColor('#1A1A1A');
text.setBold(false);
text.setItalic(false);
text.setUnderline(false);

// Style a specific range (start index, end index)
text.setBold(0, 4, true); // bold first 5 characters
text.setForegroundColor(5, 10, '#2E6DA4'); // color chars 5–10
```

### Paragraph Attributes

```javascript
para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

para.setSpacingBefore(12); // points
para.setSpacingAfter(6);
para.setLineSpacing(1.2);
para.setIndentStart(18);   // left indent in points (18 pt = 0.25 in)
para.setIndentEnd(18);     // right indent
```

### Table Cell Styling

```javascript
const cell = table.getCell(rowIndex, colIndex);

cell.setBackgroundColor('#1B3A6B');
cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
cell.setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
cell.setPaddingTop(6);
cell.setPaddingBottom(6);
cell.setPaddingLeft(8);
cell.setPaddingRight(8);

// Style text inside a cell
const cellText = cell.editAsText();
cellText.setFontSize(10);
cellText.setBold(true);
cellText.setForegroundColor('#FFFFFF');
```

### Setting Document Metadata

```javascript
doc.setName('Proposal — Client Name — March 2026');
// Note: DocumentApp does not expose author/description metadata directly.
// Use DriveApp for that:
const file = DriveApp.getFileById(doc.getId());
file.setDescription('Proposal generated by TAK Scripts');
```

### Full Document Generation Pattern

```javascript
function generateDocument(data) {
  const doc = DocumentApp.create(data.title);
  const body = doc.getBody();
  body.clear();

  setPageSetup(doc);       // margins
  buildCoverPage(body, data);
  appendTableOfContents(body);
  appendSection(body, 'Executive Summary', data.summary);
  appendSection(body, 'Scope of Work', data.scope);
  appendPricingTable(body, data.lineItems);
  appendSignatureBlock(body, data);
  addHeaderFooter(doc, data.title, data.docNumber);

  doc.saveAndClose();
  return doc.getUrl();
}
```

### Common Pitfalls in Apps Script

- **`appendParagraph` vs `insertParagraph`:** `appendParagraph` adds to the end. To insert at a specific position, use `body.insertParagraph(index, text)`. Track your insertion index carefully.
- **Heading style resets:** When you call `setHeading()`, it may reset font properties to Docs defaults. Always apply font/color/size **after** calling `setHeading()`.
- **Table borders:** `table.setBorderWidth(0)` removes the outer border. Individual cell border control (inner borders only) requires the Docs REST API — `DocumentApp` cannot set per-side borders on cells.
- **Image blobs:** Images must be fetched as a `Blob` (e.g., `UrlFetchApp.fetch(url).getBlob()`) before inserting. Store frequently used images (like your logo) in Google Drive and fetch via `DriveApp.getFileById(LOGO_FILE_ID).getBlob()`.
- **Execution time limits:** Complex documents with many tables can approach the 6-minute Apps Script execution limit. Batch your writes, avoid redundant `getBody()` calls, and consider breaking generation across multiple functions.
- **`doc.saveAndClose()`:** Always call this at the end. Without it, some changes may not persist, especially when the document URL is shared immediately after generation.

---

*Reference version: March 2026. Applies to Google Workspace (Docs, Apps Script). All hex colors are sRGB. Font availability subject to Google Fonts library updates.*
