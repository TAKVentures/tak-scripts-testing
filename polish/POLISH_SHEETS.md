# Google Sheets Dashboard Design Reference — POLISH_SHEETS.md

A practical design reference for polishing Google Apps Scripts that create or modify Sheets dashboards and data views. Every recommendation here is specific and actionable.

---

## Table of Contents

1. [Color Palettes](#1-color-palettes)
2. [Typography](#2-typography)
3. [Cell Formatting](#3-cell-formatting)
4. [Dashboard Layout](#4-dashboard-layout)
5. [Borders](#5-borders)
6. [Conditional Formatting](#6-conditional-formatting)
7. [Column Widths](#7-column-widths)
8. [Frozen Rows and Columns](#8-frozen-rows-and-columns)
9. [Sheet Tabs](#9-sheet-tabs)
10. [Charts](#10-charts)
11. [What to Avoid](#11-what-to-avoid)
12. [Apps Script Implementation Notes](#12-apps-script-implementation-notes)

---

## 1. Color Palettes

### Primary Palette — Blue/Slate (Professional Default)

Use this for internal ops dashboards, data views, and reports.

| Role                  | Hex       | Usage                                      |
|-----------------------|-----------|--------------------------------------------|
| Header background     | `#1A237E` | Section headers, top-level column headers  |
| Header text           | `#FFFFFF` | Text on dark header backgrounds            |
| Subheader background  | `#3949AB` | Secondary headers, group dividers          |
| Subheader text        | `#FFFFFF` | Text on subheader backgrounds              |
| Accent                | `#1565C0` | Stat card borders, highlighted sections    |
| Alternating row A     | `#FFFFFF` | Odd rows (white)                           |
| Alternating row B     | `#F3F6FB` | Even rows (very light blue-gray)           |
| Grid lines            | `#CFD8DC` | Inner cell borders                         |
| Neutral background    | `#ECEFF1` | Sheet canvas background, unused areas      |

### Secondary Palette — Teal/Green (Operational / TAK-style)

Use for field operations, tracking sheets, or customer-facing views.

| Role                  | Hex       | Usage                                      |
|-----------------------|-----------|--------------------------------------------|
| Header background     | `#004D40` | Section headers                            |
| Header text           | `#FFFFFF` | Text on dark backgrounds                   |
| Subheader background  | `#00796B` | Secondary headers                          |
| Accent                | `#00BFA5` | Highlights, active status indicators       |
| Alternating row A     | `#FFFFFF` | Odd rows                                   |
| Alternating row B     | `#F1F8F7` | Even rows (very light teal-gray)           |
| Grid lines            | `#B2DFDB` | Inner borders                              |
| Neutral background    | `#E0F2F1` | Canvas/padding areas                       |

### Status Badge Colors

Consistent across all sheets. Never deviate from these for status columns.

| Status              | Background | Text      | Notes                                      |
|---------------------|------------|-----------|--------------------------------------------|
| Active / OK         | `#E8F5E9` | `#1B5E20` | Light green bg, dark green text            |
| Warning / Pending   | `#FFF8E1` | `#E65100` | Light amber bg, dark orange text           |
| Error / Inactive    | `#FFEBEE` | `#B71C1C` | Light red bg, dark red text                |
| Info / Neutral      | `#E3F2FD` | `#0D47A1` | Light blue bg, dark blue text              |
| Archived / Closed   | `#F5F5F5` | `#616161` | Light gray bg, medium gray text            |

**Do not use saturated/solid red or green fills for entire rows.** Use the light badge colors above — solid fills are visually noisy and make text harder to read at a glance.

### Stat Card Colors (Row 1 Dashboards)

Stat cards should have a colored left border accent only, not a full colored background.

| Card Type      | Left Border Color | Background |
|----------------|-------------------|------------|
| Primary metric | `#1565C0`         | `#FFFFFF`  |
| Success metric | `#2E7D32`         | `#FFFFFF`  |
| Warning metric | `#F57F17`         | `#FFFFFF`  |
| Danger metric  | `#C62828`         | `#FFFFFF`  |
| Neutral metric | `#455A64`         | `#F8F9FA`  |

---

## 2. Typography

### Font Choices

Google Sheets supports a limited font set. Use only these:

- **Roboto** — Default preferred font. Clean, modern, excellent legibility at all sizes. Use for all data views.
- **Arial** — Fallback when Roboto feels too casual. Slightly more "corporate." Use for financial/exec dashboards.
- **Roboto Mono** — For numeric-heavy columns (IDs, codes, coordinates). Monospaced numerals align better.

Never use: Comic Sans, Impact, Times New Roman, or decorative fonts in any professional sheet.

### Font Size Reference

| Element                  | Size | Style           | Font        |
|--------------------------|------|-----------------|-------------|
| Sheet title (A1)         | 18pt | Bold            | Roboto      |
| Section header           | 12pt | Bold            | Roboto      |
| Column header (row 3)    | 10pt | Bold            | Roboto      |
| Data cells               | 10pt | Regular         | Roboto      |
| Sub-labels / helper text | 9pt  | Regular, Italic | Roboto      |
| Stat card value          | 20pt | Bold            | Roboto      |
| Stat card label          | 9pt  | Regular         | Roboto      |
| Footer / timestamp       | 8pt  | Italic          | Arial       |

### Bold and Italic Usage Rules

- **Bold** is reserved for: headers, column labels, stat card values, and total/summary rows.
- *Italic* is reserved for: timestamps, helper text, footnotes, and formula-derived labels.
- Never bold entire data rows. Only bold a data cell if it is a summary (Total, Average, Max).
- Do not mix bold and italic in the same cell unless it is a very intentional emphasis (rare).

---

## 3. Cell Formatting

### Row Height (Padding)

Row height is the single biggest legibility improvement you can make. Cramped rows feel cheap.

| Row Type              | Height (px) | Notes                                  |
|-----------------------|-------------|----------------------------------------|
| Header row (titles)   | 40          | Generous padding for visual weight     |
| Column header row     | 30          | Clear separation from data             |
| Data rows             | 24          | Default comfortable reading height     |
| Stat card rows        | 60–80       | Tall enough to display large numbers   |
| Footer / timestamp    | 20          | Compact, secondary importance          |

The default row height in Sheets is 21px. Always set it explicitly — never rely on defaults.

### Text Alignment by Data Type

Alignment is not aesthetic preference — it communicates data type to the reader.

| Data Type             | Horizontal | Vertical | Notes                                  |
|-----------------------|------------|----------|----------------------------------------|
| Column headers        | Center     | Middle   | Always centered, always middle         |
| Text / Names          | Left       | Middle   | Never center-align text columns        |
| Numbers / Currency    | Right      | Middle   | Right-align to stack decimal points    |
| Dates                 | Center     | Middle   | Dates are identifiers, not prose       |
| Status badges         | Center     | Middle   | Short fixed strings, center works      |
| IDs / Codes           | Center     | Middle   | Fixed-width, centering looks clean     |
| Email addresses       | Left       | Middle   | Long strings need left alignment       |
| Percentages           | Right      | Middle   | Same rule as numbers                   |
| Checkboxes / Booleans | Center     | Middle   | Visual symbol, always center           |

### Number Formats

Always explicitly set number formats. Never rely on Sheets auto-detection.

| Data Type           | Format String              | Example Output      |
|---------------------|----------------------------|---------------------|
| Currency USD        | `"$"#,##0.00`              | $1,234.56           |
| Currency (no cents) | `"$"#,##0`                 | $1,235              |
| Percentage          | `0.0%`                     | 84.3%               |
| Integer             | `#,##0`                    | 1,234               |
| Decimal 2           | `#,##0.00`                 | 1,234.56            |
| Date                | `MM/DD/YYYY`               | 03/23/2026          |
| Date + Time         | `MM/DD/YYYY HH:mm`         | 03/23/2026 14:30    |
| Duration            | `[h]:mm:ss`                | 2:45:00             |
| Phone number        | `"("###") "###"-"####`     | (555) 123-4567      |
| ID / Code           | `@` (plain text)           | TA-00123            |

---

## 4. Dashboard Layout

### Standard Layout Grid

Follow this row allocation pattern for every dashboard sheet.

```
Row 1:    [Stat Card 1]  [Stat Card 2]  [Stat Card 3]  [Stat Card 4]
Row 2:    [Label 1]      [Label 2]      [Label 3]      [Label 4]
Row 3:    [BLANK — visual separator, height 8px]
Row 4:    [Column Header 1] [Column Header 2] ... [Column Header N]
Rows 5+:  [Data Row]
Last Row: [Totals / Summary Row]
```

### Stat Card Design (Rows 1–2)

Each stat card occupies multiple merged columns. Typical layout for a 12-column sheet:

- Columns A–C: Stat Card 1
- Columns D–F: Stat Card 2
- Columns G–I: Stat Card 3
- Columns J–L: Stat Card 4

Row 1 contains the large numeric/metric value (20pt bold).
Row 2 contains the descriptive label (9pt regular, gray `#757575`).

**Never put both the value and label in the same merged cell.** Keeping them in separate rows lets you format them independently and makes future edits clean.

### Section Headers

When a sheet has multiple data sections (not just one table), separate them with section headers:

```
[SECTION HEADER — spans full width, colored background, bold, 12pt]
[Column Headers]
[Data Rows]
[Blank separator row — height 8px, background #ECEFF1]
[Next SECTION HEADER]
```

### Column Header Row Design

- Background: `#1A237E` (or your primary header color)
- Text: `#FFFFFF`
- Font: 10pt Bold Roboto
- Height: 30px
- Text wrap: Off (truncate, not wrap — headers should be short)
- Freeze this row always

### Data Row Design

- Alternating background: rows alternate between `#FFFFFF` and `#F3F6FB`
- Text: `#212121` (near-black, not pure black — softer on eyes)
- Font: 10pt Regular Roboto
- Height: 24px
- Text wrap: Off unless the column is explicitly a "notes" or "description" column

### Summary / Totals Row

- Background: `#E8EAF6` (light indigo)
- Text: `#1A237E` (primary header color)
- Font: 10pt Bold
- Top border: 2px solid `#3949AB`

---

## 5. Borders

### Border Philosophy

Less is more. Borders exist to guide the eye, not to cage every cell.

### Border Hierarchy

| Context                        | Style          | Color     | Thickness |
|--------------------------------|----------------|-----------|-----------|
| Outer sheet border             | Solid          | `#455A64` | 2px       |
| Section separator              | Solid          | `#3949AB` | 2px       |
| Column header bottom           | Solid          | `#1A237E` | 2px       |
| Data row inner (horizontal)    | Solid          | `#E0E0E0` | 1px       |
| Data row inner (vertical)      | None or dotted | `#E0E0E0` | 1px       |
| Summary row top                | Solid          | `#3949AB` | 2px       |
| Stat card outline              | Solid          | `#CFD8DC` | 1px       |

### When NOT to Use Borders

- **No vertical inner borders** on data rows unless columns are closely related and you need visual separation (e.g., a multi-column "address" group).
- **No borders on blank/spacer rows.** Leave them clean.
- **No borders on the sheet canvas** outside your data range.
- **No borders around individual stat card metric cells** — use the card outline border only.

### Apps Script Border Constants

```javascript
// Border styles available in Apps Script
SpreadsheetApp.BorderStyle.SOLID        // thin solid (1px)
SpreadsheetApp.BorderStyle.SOLID_MEDIUM // medium solid (2px)
SpreadsheetApp.BorderStyle.SOLID_THICK  // thick solid (3px)
SpreadsheetApp.BorderStyle.DASHED
SpreadsheetApp.BorderStyle.DOTTED
SpreadsheetApp.BorderStyle.DOUBLE

// setBorder(top, left, bottom, right, vertical, horizontal, color, style)
range.setBorder(true, true, true, true, false, false,
  '#455A64', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
```

---

## 6. Conditional Formatting

### Status Column Rules

Apply in this exact priority order (higher priority rules listed first):

```
Rule 1: Text is exactly "Error"    → bg #FFEBEE, text #B71C1C
Rule 2: Text is exactly "Inactive" → bg #FFEBEE, text #B71C1C
Rule 3: Text is exactly "Warning"  → bg #FFF8E1, text #E65100
Rule 4: Text is exactly "Pending"  → bg #FFF8E1, text #E65100
Rule 5: Text is exactly "Active"   → bg #E8F5E9, text #1B5E20
Rule 6: Text is exactly "Complete" → bg #E8F5E9, text #1B5E20
```

### Numeric Threshold Rules

For percentage columns (e.g., completion rate, capacity):

```
< 50%:  bg #FFEBEE, text #B71C1C  (critical)
50–74%: bg #FFF8E1, text #E65100  (warning)
75–89%: bg #FFFDE7, text #827717  (caution)
90%+:   bg #E8F5E9, text #1B5E20  (good)
```

For date columns (overdue detection):

```
Date < TODAY():  bg #FFEBEE, text #B71C1C  (overdue)
Date = TODAY():  bg #FFF8E1, text #E65100  (due today)
Date > TODAY():  No formatting              (upcoming — keep clean)
```

### Conditional Formatting Best Practices

- Apply conditional formatting **only to specific columns** — never to full rows unless the entire row's status depends on a single cell and you've deliberately made that UX choice.
- Keep the rule count under 10 per sheet. More than that degrades performance and becomes unmaintainable.
- Use **Custom formula** rules (`=` prefix) when you need cross-column logic. Example: highlight a row if column C equals "Error": `=$C2="Error"`.
- Always test conditional formatting rules after applying banding (alternating colors) — they stack in unexpected ways.

### Apps Script Conditional Formatting

```javascript
function applyStatusConditionalFormatting(sheet, statusColumnRange) {
  const rules = [];

  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Error')
    .setBackground('#FFEBEE')
    .setFontColor('#B71C1C')
    .setRanges([statusColumnRange])
    .build();

  const warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Warning')
    .setBackground('#FFF8E1')
    .setFontColor('#E65100')
    .setRanges([statusColumnRange])
    .build();

  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Active')
    .setBackground('#E8F5E9')
    .setFontColor('#1B5E20')
    .setRanges([statusColumnRange])
    .build();

  rules.push(errorRule, warningRule, activeRule);
  sheet.setConditionalFormatRules(rules);
}
```

---

## 7. Column Widths

### Standard Width Reference

These are battle-tested widths. Use them as your starting point.

| Content Type              | Width (px) | Notes                                      |
|---------------------------|------------|--------------------------------------------|
| Row number / index        | 40         | Just wide enough for 4-digit numbers       |
| Checkbox / boolean        | 50         | Minimal — just the checkbox widget         |
| Status badge              | 90         | Fits most status strings + padding         |
| Short ID / code           | 80         | 6–8 character codes                        |
| Long ID / UUID            | 220        | Full UUID display                          |
| First name                | 100        | Single name field                          |
| Full name                 | 160        | First + Last                               |
| Email address             | 220        | Most emails fit; clip long ones            |
| Phone number              | 120        | Formatted phone number                     |
| Date only                 | 100        | MM/DD/YYYY                                 |
| Date + time               | 140        | MM/DD/YYYY HH:mm                           |
| Short number (0–999)      | 70         | Counts, small quantities                   |
| Currency / decimal        | 100        | Room for $X,XXX.XX                         |
| Percentage                | 70         | 3-character display: 84%                   |
| Short text / category     | 120        | Tags, types, short enums                   |
| Medium text / description | 200        | Truncated notes, addresses                 |
| Long text / notes         | 300        | Multi-word descriptions (wrap enabled)     |
| URL                       | 200        | Display with hyperlink, truncate           |
| Icon / emoji column       | 40         | Just the symbol                            |

### Width Setting in Apps Script

```javascript
// Set individual column width
sheet.setColumnWidth(columnIndex, widthInPixels);  // columnIndex is 1-based

// Set multiple columns to the same width
sheet.setColumnWidths(startColumn, numColumns, widthInPixels);

// Example: set columns 1–5 to 120px, column 6 to 200px
sheet.setColumnWidths(1, 5, 120);
sheet.setColumnWidth(6, 200);
```

---

## 8. Frozen Rows and Columns

### When to Freeze

**Always freeze:**
- Row 1 if it contains column headers (single-table sheets)
- Row 4 if using the dashboard layout (rows 1–3 are stat cards + spacer, row 4 is headers)
- The first column if it contains a primary identifier (name, ID) that anchors each row

**Freeze more aggressively when:**
- The sheet has more than 15 columns (freeze identifier + 1–2 context columns)
- Users will scroll horizontally in normal workflow
- The sheet is used on mobile or tablet (freeze 1 column minimum)

**Do not freeze:**
- More than 3 rows — it wastes vertical screen real estate
- More than 2 columns on narrow datasets — it causes awkward proportions
- Any row that is a spacer or decorative element

### Standard Freeze Configurations

| Sheet Type                    | Frozen Rows | Frozen Columns |
|-------------------------------|-------------|----------------|
| Simple data table             | 1           | 0              |
| Data table with ID column     | 1           | 1              |
| Dashboard with stat cards     | 4           | 0              |
| Wide operational table        | 1           | 2              |
| Lookup / reference table      | 1           | 1              |

### Apps Script Implementation

```javascript
// Freeze rows and columns
sheet.setFrozenRows(1);
sheet.setFrozenColumns(1);

// For dashboard layout with 4 header rows
sheet.setFrozenRows(4);
sheet.setFrozenColumns(0); // no column freeze on dashboard
```

---

## 9. Sheet Tabs

### Naming Conventions

Tab names must be short, descriptive, and scannable. Rules:

- **Max 20 characters** — longer names truncate in the tab bar
- **Title Case** — not ALL CAPS, not lowercase
- **No special characters** except spaces and hyphens
- **Lead with the content type** when ordering matters: "Members", "Events", "Archive"

| Tab Type               | Naming Pattern          | Example               |
|------------------------|-------------------------|-----------------------|
| Primary data view      | Noun (plural)           | Members, Events, Tasks|
| Dashboard / summary    | "Dashboard" suffix      | Team Dashboard        |
| Config / settings      | "Config" or "Settings"  | Sheet Config          |
| Lookup / reference     | "Ref -" prefix          | Ref - Statuses        |
| Archive                | "Archive -" prefix      | Archive - 2025        |
| Staging / temp         | "TEMP -" prefix         | TEMP - Import         |
| Log / audit            | "Log" suffix            | Change Log            |
| Instructions           | "README" or "Guide"     | README                |

### Tab Colors

Use tab colors to communicate category, not status. Keep it to 4–5 colors maximum across the entire workbook.

| Tab Category           | Tab Color  | Hex       |
|------------------------|------------|-----------|
| Primary / live data    | Blue       | `#1565C0` |
| Dashboard / summary    | Dark green | `#2E7D32` |
| Config / settings      | Gray       | `#546E7A` |
| Reference data         | Teal       | `#00695C` |
| Archive                | Brown      | `#4E342E` |
| Staging / temp         | Orange     | `#E65100` |
| README / docs          | Purple     | `#4527A0` |

### Tab Ordering

Order tabs left to right in the sequence a user would naturally work:

```
README → Dashboard → [Primary Data Tabs] → [Reference Tabs] → Config → Archive
```

Never mix live data tabs and archive tabs. Archives always go rightmost.

### Apps Script Tab Management

```javascript
// Set tab color
sheet.setTabColor('#1565C0');

// Rename sheet
sheet.setName('Members');

// Move sheet to a specific position (0-indexed)
spreadsheet.setActiveSheet(sheet);
spreadsheet.moveActiveSheet(1);  // move to second position (after README)
```

---

## 10. Charts

### When to Use Charts

Only add charts if:
- The data has a meaningful trend over time, OR
- You need to show proportional breakdown (composition)
- The chart will be reviewed regularly, not just at sheet creation

Do not add charts to every dashboard by default. A clean stat card layout often communicates faster than a chart.

### Chart Style Rules

**Preferred chart types:**
- **Line chart** — time-series trends (disable smooth curves)
- **Column chart** — comparisons across categories
- **Stacked bar** — composition over time or across groups
- **Pie/donut** — only when showing 2–4 segments; any more, use a bar chart instead

**Formatting rules:**
- Background: `#FFFFFF` (white, never transparent in dashboards)
- Chart border: 1px `#CFD8DC`
- No chart title if the surrounding dashboard context makes it obvious
- Legend: bottom position, 9pt Roboto
- Gridlines: light gray `#E0E0E0`, horizontal only
- No 3D effects, no shadows, no rounded bars
- Remove all gradient fills — use flat solid colors only
- Data label font: 9pt, same color as the series it labels

**Color sequence for chart series (in order):**

1. `#1565C0` (blue)
2. `#2E7D32` (green)
3. `#E65100` (orange)
4. `#6A1B9A` (purple)
5. `#00838F` (teal)

Never use red as a series color unless it explicitly represents an error/danger metric — readers will assume something is wrong.

### Apps Script Chart Insertion

```javascript
function insertCleanLineChart(sheet, dataRange, position) {
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setPosition(position.row, position.col, 0, 0)
    .setOption('title', '')
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('legend.position', 'bottom')
    .setOption('legend.textStyle.fontSize', 9)
    .setOption('hAxis.gridlines.color', 'transparent')
    .setOption('vAxis.gridlines.color', '#E0E0E0')
    .setOption('chartArea.width', '85%')
    .setOption('chartArea.height', '75%')
    .setOption('colors', ['#1565C0', '#2E7D32', '#E65100']);

  sheet.insertChart(chartBuilder.build());
}
```

---

## 11. What to Avoid

### Color Mistakes

- **Too many colors.** If your sheet uses more than 5 distinct background colors, you have too many. Each additional color reduces the signal value of all others.
- **Saturated solid fills on data rows.** Solid `#FF0000` or `#00FF00` rows are jarring. Always use the light badge variants from Section 1.
- **Random accent colors.** Every color in the sheet should map to a meaning. If a color doesn't mean anything specific, remove it.
- **Low-contrast combinations.** Yellow text on white, light gray text on white, red text on dark red. Verify contrast — aim for WCAG AA minimum (4.5:1 ratio for normal text).
- **Gradient fills.** Sheets supports gradient cell fills. Do not use them. They look unprofessional in data contexts.

### Typography Mistakes

- **Multiple font families in the same sheet.** Pick one: Roboto or Arial. Not both.
- **Font sizes below 8pt.** Sheets renders these poorly and they are unreadable on most monitors.
- **All-caps data cells.** ALL CAPS for headers is acceptable but for data cells it reads as aggressive and degrades scan speed.
- **Font size inflation.** Making column headers 14pt does not make them more important — it makes the sheet feel amateurish. Stick to the size table in Section 2.
- **Italic on column headers.** Column headers should always be bold regular, never italic.

### Layout Mistakes

- **Merged cells in data ranges.** Merging cells breaks sorting, filtering, copy-paste, and formula ranges. Only merge cells in header/title rows that will never be part of a data operation.
- **Empty columns as spacers.** Instead of inserting blank column B to create visual padding, use proper column widths and alignment.
- **Data starting in row 1.** Always leave row 1 for a title or stat summary. Data starting at A1 with no header row is a maintenance and usability problem.
- **No frozen rows.** If someone scrolls down 200 rows and cannot see column headers, the sheet fails at its job.
- **Inconsistent row heights.** If row 5 is 21px and row 6 is 24px, it looks broken. Always set heights explicitly and uniformly.

### Performance Mistakes

- **Applying formatting cell by cell in a loop.** This is the most common Apps Script performance killer. Batch all formatting calls (see Section 12).
- **Too many conditional formatting rules.** More than 20 rules on a sheet causes noticeable lag. Consolidate with custom formula rules.
- **Setting format on entire columns.** `sheet.getRange('A:A').setBackground(...)` applies to over 10,000 rows. Only format the rows that contain data.
- **Charts on sheets with heavy conditional formatting.** Both compete for rendering resources. Consider a separate "Charts" sheet.

### UX Mistakes

- **Hyperlinks with no visual cue.** If a cell contains a link, give it the standard `#1565C0` color and underline so users know it is clickable.
- **Notes/comments as primary data.** Cell notes are invisible until hovered. Never put data that matters in a cell note.
- **Hiding columns instead of deleting them.** Hidden columns confuse collaborators and break formula debugging. Delete what you do not need, or move it to a reference tab.
- **No sheet protection on header rows.** Dashboard headers and stat card rows should be set to warning-only protection so accidental edits are flagged.

---

## 12. Apps Script Implementation Notes

### Batch API Call Patterns

**The cardinal rule: minimize calls to the Sheets API.** Every `getValue()`, `setBackground()`, or `getRange()` call on an individual cell is a round trip. Batch everything.

```javascript
// BAD — individual calls in a loop (extremely slow)
for (let i = 0; i < data.length; i++) {
  sheet.getRange(i + 2, 1).setValue(data[i].name);
  sheet.getRange(i + 2, 1).setBackground('#F3F6FB');
  sheet.getRange(i + 2, 1).setFontSize(10);
}

// GOOD — batch writes, then batch format
// Step 1: Write all data in one call
const outputData = data.map(row => [row.name, row.status, row.date]);
sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);

// Step 2: Format the entire data range in one chained call
const dataRange = sheet.getRange(2, 1, outputData.length, outputData[0].length);
dataRange
  .setFontFamily('Roboto')
  .setFontSize(10)
  .setFontColor('#212121')
  .setVerticalAlignment('middle');

// Step 3: Apply alternating row colors with banding (more efficient than per-row calls)
const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
banding.setHeaderRowColor('#1A237E');
banding.setFirstRowColor('#FFFFFF');
banding.setSecondRowColor('#F3F6FB');
```

### Complete Dashboard Setup Function Pattern

```javascript
function formatDashboardSheet(sheet, options = {}) {
  const {
    headerColor      = '#1A237E',
    headerTextColor  = '#FFFFFF',
    dataRowColor1    = '#FFFFFF',
    dataRowColor2    = '#F3F6FB',
    headerRowIndex   = 4,
    dataStartRow     = 5,
    numDataRows      = 0,
    numCols          = 10,
    rowHeightHeader  = 30,
    rowHeightData    = 24,
    rowHeightStatVal = 60,
    rowHeightStatLbl = 24,
    freezeRows       = 4,
    freezeCols       = 0,
  } = options;

  // --- Stat card rows ---
  sheet.setRowHeight(1, rowHeightStatVal);
  sheet.setRowHeight(2, rowHeightStatLbl);
  sheet.setRowHeight(3, 8);  // spacer

  // --- Column header row ---
  sheet.setRowHeight(headerRowIndex, rowHeightHeader);
  const headerRange = sheet.getRange(headerRowIndex, 1, 1, numCols);
  headerRange
    .setBackground(headerColor)
    .setFontColor(headerTextColor)
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // --- Data rows ---
  if (numDataRows > 0) {
    sheet.setRowHeights(dataStartRow, numDataRows, rowHeightData);
    const dataRange = sheet.getRange(dataStartRow, 1, numDataRows, numCols);
    dataRange
      .setFontFamily('Roboto')
      .setFontSize(10)
      .setFontColor('#212121')
      .setVerticalAlignment('middle');

    // Alternating row banding
    const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding.setFirstRowColor(dataRowColor1);
    banding.setSecondRowColor(dataRowColor2);
  }

  // --- Freeze rows and columns ---
  sheet.setFrozenRows(freezeRows);
  sheet.setFrozenColumns(freezeCols);

  // --- Header bottom border ---
  headerRange.setBorder(
    false, false, true, false, false, false,
    headerColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // --- Spacer row background ---
  sheet.getRange(3, 1, 1, numCols).setBackground('#ECEFF1');
}
```

### Key Formatting Method Reference

```javascript
const range = sheet.getRange('A1:D10');

// Background and text color
range.setBackground('#1A237E');
range.setFontColor('#FFFFFF');

// Font properties
range.setFontFamily('Roboto');
range.setFontSize(10);
range.setFontWeight('bold');    // 'bold' or 'normal'
range.setFontStyle('italic');   // 'italic' or 'normal'

// Alignment
range.setHorizontalAlignment('center');  // 'left', 'center', 'right'
range.setVerticalAlignment('middle');    // 'top', 'middle', 'bottom'

// Text wrap
range.setWrap(false);  // false = clip/truncate (preferred for data)
range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

// Number format
range.setNumberFormat('"$"#,##0.00');
range.setNumberFormat('MM/DD/YYYY');
range.setNumberFormat('0.0%');

// Row and column sizing
sheet.setRowHeight(rowIndex, 24);              // single row
sheet.setRowHeights(startRow, numRows, 24);    // range of rows
sheet.setColumnWidth(colIndex, 120);           // single column
sheet.setColumnWidths(startCol, numCols, 120); // range of columns

// Borders — setBorder(top, left, bottom, right, vertical, horizontal, color, style)
range.setBorder(
  true, true, true, true, false, false,
  '#455A64', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
);

// Freeze
sheet.setFrozenRows(1);
sheet.setFrozenColumns(1);

// Tab color
sheet.setTabColor('#1565C0');

// Merge cells (headers only — never in data ranges)
range.merge();
range.mergeHorizontally();
range.unmerge();

// Banding (alternating rows — more efficient than per-row coloring)
const banding = range.applyRowBanding();
banding.setFirstRowColor('#FFFFFF');
banding.setSecondRowColor('#F3F6FB');
banding.setHeaderRowColor('#1A237E');
banding.setFooterRowColor('#E8EAF6');
```

### SpreadsheetApp.flush() — When to Use

`SpreadsheetApp.flush()` forces all pending changes to be written immediately. Use it:
- After a batch of writes before reading values that depend on those writes
- At the end of long-running functions to ensure partial progress is saved
- Before `Utilities.sleep()` calls if doing rate-limited work

Do not call `flush()` inside loops — it negates all batching gains.

```javascript
// Pattern: write all → flush once → then format
sheet.getRange(2, 1, data.length, cols).setValues(data);
SpreadsheetApp.flush();  // ensure writes complete before formatting pass
const dataRange = sheet.getRange(2, 1, data.length, cols);
dataRange.setFontFamily('Roboto').setFontSize(10);
```

### Clearing Before Re-applying

When a script rebuilds a sheet, always clear old formatting first to prevent stale styles from conflicting with new ones:

```javascript
const fullRange = sheet.getDataRange();
fullRange.clearFormat();   // clears formatting only, keeps values
// or
fullRange.clearContent();  // clears values only, keeps formatting
// or
sheet.clear();             // clears everything including formatting and notes

// Also remove any existing banding before re-applying
const bandings = fullRange.getBandings();
bandings.forEach(b => b.remove());
```

### Protecting Header Rows

For dashboards where users should not accidentally overwrite headers or stat cards:

```javascript
function protectHeaderRows(sheet) {
  const protection = sheet.getRange('1:4').protect();
  protection.setDescription('Dashboard headers — do not edit');
  protection.setWarningOnly(true);  // warns but does not fully lock
}
```

### Hyperlink Formatting

Links in cells should always be visually distinct:

```javascript
// Set a cell to display as a hyperlink with standard link color
const linkRange = sheet.getRange(row, col);
linkRange.setFormula(`=HYPERLINK("${url}", "${displayText}")`);
linkRange.setFontColor('#1565C0');
linkRange.setFontLine('underline');
```

---

*End of POLISH_SHEETS.md — Last updated 2026-03-23*
