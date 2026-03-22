# TAKScripts Design Guide

Read this before testing any script. Use these standards when evaluating the visual quality and customer experience of each script.

---

## Brand Identity

- **Brand:** TAKScripts by TAK Ventures
- **Tagline:** "Sometimes all you need is a little digital help."
- **Logo:** 🕷 Spider icon
- **Domain:** takscripts.store

## Color Palette

| Purpose | Color | Hex |
|---------|-------|-----|
| Primary Gold | Gold | #C9A84C |
| Dark Background | Near Black | #1A1A1A |
| Header Text | White | #FFFFFF |
| Body Text | Dark Gray | #333333 |
| Muted Text | Medium Gray | #666666 |
| Light Text | Light Gray | #999999 |
| Success | Green | #2E7D32 |
| Error | Red | #C62828 |
| Warning | Amber | #F57F17 |
| Info | Blue | #1565C0 |
| Sheet Header BG | Near Black | #1A1A1A |
| Sheet Header Text | Gold | #C9A84C |
| Sheet Accent Row | Light Gold BG | #FFF8E7 |
| Sheet Border | Light Gray | #E0E0E0 |

## Google Sheets Design Standards

Every script that creates or modifies a Google Sheet should follow these rules:

### Header Row
- Background: `#1A1A1A` (near black)
- Text color: `#C9A84C` (gold)
- Font: `Roboto Mono`, Bold, 10pt
- Frozen (always visible when scrolling)
- Row height: 32px

### Column Widths
- Timestamps: 180px
- Email addresses: 250px
- Names: 150px
- Descriptions/subjects: 300px
- Status columns: 100px
- Always auto-resize if content exceeds width

### Data Rows
- Font: `Roboto`, Regular, 10pt
- Alternating row colors: White `#FFFFFF` / Light gray `#F9F9F9`
- Vertical alignment: Middle
- Horizontal alignment: Left (text), Center (status/dates)

### Status Cells
Use colored backgrounds for status indicators:
- Active/Success: Light green `#E8F5E9` with text `#2E7D32`
- Pending/Warning: Light amber `#FFF8E1` with text `#F57F17`
- Error/Failed: Light red `#FFEBEE` with text `#C62828`
- Completed: Light blue `#E3F2FD` with text `#1565C0`

### Sheet Tab Naming
- Use descriptive names with emoji prefix: "📊 Sales Log", "📬 Reply Log", "⚙️ Settings"
- Never leave default "Sheet1" names

### Footer Row (optional)
- Add a subtle footer in the last used row + 2:
- "Powered by TAKScripts · takscripts.store"
- Color: `#CCCCCC`, Font: 9pt, Italic

---

## Gmail / Email Design Standards

When scripts send emails (auto-replies, notifications, reports):

### Email Body
- Use clean HTML formatting, not plain text
- Include a subtle header: "🕷 TAKScripts" in gold
- Body font: system font stack, 14px, line-height 1.6
- Sign off with: "— Sent by [Script Name] via TAKScripts"
- Include a subtle footer link: "Manage this automation at takscripts.store"

### Email Subject Lines
- Clear and professional
- Include a prefix when appropriate: "[OOO]", "[Report]", "[Alert]"
- Never use ALL CAPS

---

## Google Drive Design Standards

When scripts create folders or files:

### Folder Naming
- Use emoji prefix: "📁 Client Files", "📊 Reports"
- Consistent naming: "TAKScripts — [Purpose]"

### File Naming
- Descriptive: "Invoice_Log_2026-03-22.pdf"
- Include dates where relevant
- Never use spaces — use underscores

---

## Sidebar / Dialog UI Standards

When scripts use HtmlService for custom UI:

### Layout
- Dark header bar: `#1A1A1A` background with 🕷 logo and gold "TAKScripts" text
- White/light body: `#FAFAFA` background
- Clear section separation with subtle borders
- Consistent 16px padding

### Form Elements
- Rounded inputs (6px border-radius)
- Gold focus state: `border-color: #C9A84C` with `box-shadow: 0 0 0 3px rgba(201,168,76,0.1)`
- Uppercase labels: 11px, `letter-spacing: 1px`, `#666666`
- Help text below inputs: 11px, `#999999`

### Buttons
- Primary: Dark bg `#1A1A1A`, gold text `#C9A84C`, gold border
- Primary hover: Gold bg, dark text
- Secondary: White bg, gray text, gray border
- Full width, 12px padding, 8px border-radius
- 600 font weight, 0.5px letter spacing

### Feedback
- Success message: Green background `#E8F5E9`, green text
- Error message: Red background `#FFEBEE`, red text
- Always show feedback after user actions (save, submit, etc.)

---

## Custom Menu Standards

Every script should add a menu to the Google app it runs in:

### Menu Name
- Always: "🕷 TAKScripts"
- If multiple scripts installed, use submenus: "🕷 TAKScripts" → "OOO Auto-Reply" → actions

### Menu Items
- Use emoji prefixes for clarity:
  - ⚙️ Settings
  - ▶️ Start / Run
  - ⏹ Stop
  - 🧪 Test Run
  - 📊 View Log / Report
  - ℹ️ About TAKScripts
- Always include a separator before "About"
- Always include a "Test Run" option

---

## Quality Checklist for Design Review

When testing, check ALL of these visual/UX elements:

- [ ] Sheet headers use brand colors (dark bg, gold text)
- [ ] Sheet has frozen header row
- [ ] Column widths are appropriate (no truncated text)
- [ ] Alternating row colors are applied
- [ ] Status cells use color-coded backgrounds
- [ ] Custom menu appears with proper emoji icons
- [ ] Sidebar/dialog has branded header
- [ ] Form inputs have proper focus states
- [ ] Success/error feedback is shown to user
- [ ] Emails use HTML formatting (not plain text)
- [ ] Email has branded footer
- [ ] Drive folders use emoji naming convention
- [ ] No default "Sheet1" tab names remain
- [ ] Script name and version visible somewhere in the UI
- [ ] "Powered by TAKScripts" footer where appropriate

---

*This guide is the visual standard for all TAKScripts products. Every script should feel premium, consistent, and professionally designed.*
