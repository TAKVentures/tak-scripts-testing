# TAKScripts — Script Build & Launch Workflow

The complete process for taking a script from idea to live in the store.

---

## Reference Files (read before building anything)

| File | Purpose |
|---|---|
| `BUILD_STANDARDS.md` | 14 coding rules — follow every one, no exceptions |
| `DESIGN-GUIDE.md` | Visual standards — colors, fonts, dashboard layout, sidebar style |
| `TESTING_PROTOCOL.md` | How testing works and what Browser Claude is doing |
| `TEST_TEMPLATE.md` | Blank template for writing a new TEST.md card |
| `TESTING-SHORTCUT.md` | Copy-paste prompt to hand to Browser Claude for testing |

---

## Phase 1 — Build

### Step 1: Read the standards
Before writing a single line of code, read:
- `BUILD_STANDARDS.md` — all 14 rules
- `DESIGN-GUIDE.md` — brand colors, dashboard pattern, sidebar style

### Step 2: Build the script
Write the script in:
```
/business/products/{slug}/{slug}.gs
```

Checklist while building:
- [ ] Rule 1 — PropertiesService saved before ScriptApp trigger calls
- [ ] Rule 2 — Dashboard with 3-row stats bar, `DASHBOARD_HEADER_ROW = 3`
- [ ] Rule 3 — All filter fields accept comma-separated input with OR logic
- [ ] Rule 4 — Every `getUi().alert()` in try/catch with Logger.log fallback
- [ ] Rule 5 — Processed IDs tracked in PropertiesService (if appending rows)
- [ ] Rule 6 — BRAND object copied in, no hardcoded hex values
- [ ] Rule 7 — Script header with version number
- [ ] Rule 8 — Menu order: Settings → Main Action → Dashboard → Refresh → How to Use → About
- [ ] Rule 9 — `loadSettings()` merges saved values with `getDefaultSettings_()`
- [ ] Rule 10 — Deliverable will be `.txt` (nothing to do in code, just remember)
- [ ] Rule 11 — ES5 by default unless script needs ES6
- [ ] Rule 12 — Sidebar CSS includes `textarea` in width/box-sizing selector
- [ ] Rule 13 — Save confirmation: button disables → "Saving…" → "✓ Saved!" → resets after 2.5s
- [ ] Rule 14 — `showHelp()` + `getHelpHtml_()` with Quick Start, Settings Guide, Tips sections

### Step 3: Write the SETUP.md
```
/business/products/{slug}/SETUP.md
```
Covers: what it does, installation steps (5 steps: Sheet → Extensions → paste → reload → authorize), settings guide, tips, troubleshooting. Reference any existing SETUP.md as a template.

---

## Phase 2 — Test

### Step 4: Add to the testing queue
1. Create folder: `queue/{n}-{slug}/`
2. Copy the script: `queue/{n}-{slug}/script.gs`
3. Write the test card: `queue/{n}-{slug}/TEST.md`

The TEST.md card must include:
- Pre-setup (any data to add to the sheet before testing)
- Exact settings values to enter
- The 10-step customer journey (open sheet → menu → settings → save → run → verify → run again → no duplicates → dashboard stats → help sidebar)
- Expected result for each step

Use `TEST_TEMPLATE.md` as your starting point.

### Step 5: Run the test
Hand Browser Claude the prompt from `TESTING-SHORTCUT.md`. It will:
- Read the TEST.md card
- Paste the script into the test spreadsheet
- Follow every step
- Write a `test-report.md` back to the queue folder

### Step 6: Review the report
- **Pass** → move to Phase 3
- **Fail** → fix the issue, update `script.gs` in both `business/products/` and `queue/`, re-test

---

## Phase 3 — Approve & Package

### Step 7: Move to approved
```
approved/{n}-{slug}/script.gs
```

### Step 8: Upload to Supabase Storage
```bash
# Script deliverable
PUT deliverables/{slug}/{slug}.txt

# Setup guide
PUT deliverables/{slug}/SETUP.md
```

### Step 9: Update Supabase database
In the `product_ideas` table, set:
```
deliverable_url  = 'deliverables/{slug}/{slug}.txt'
setup_guide_url  = 'deliverables/{slug}/SETUP.md'
is_active        = true
status           = 'live'
```

### Step 10: Create Stripe product
1. Create the product in Stripe dashboard
2. Copy the `price_id` (format: `price_xxxx`)
3. Set in Supabase: `stripe_price_id = 'price_xxxx'`

---

## Quick Reference — Slug & Numbering

| # | Slug | Name |
|---|---|---|
| 01 | auto-invoice-generator | Auto Invoice Generator |
| 02 | contract-proposal-generator | Contract & Proposal Generator |
| 03 | email-invoice-detector | Email Invoice Detector |
| 04 | follow-up-nudger | Follow-Up Nudger |
| 05 | inventory-low-stock-alert | Inventory Low-Stock Alert |
| 06 | vip-priority-alert | VIP Priority Alert |
| 07 | smart-ooo-auto-reply | Smart OOO Auto-Reply |
| 08 | bulk-email-unsubscriber | Bulk Email Unsubscriber |
| 09 | auto-file-organizer | Auto File Organizer |
| 10 | attachment-auto-saver | Attachment Auto-Saver |
| 11 | data-cleanup-wizard | Data Cleanup Wizard |
| 12+ | next scripts... | |

---

## Supabase Storage Structure
```
deliverables/
  {slug}/
    {slug}.txt       ← the script customers download
    SETUP.md         ← the setup guide
```

---

## Status Definitions (product_ideas table)

| Status | Meaning |
|---|---|
| `idea` | Not yet built |
| `building` | In development |
| `testing` | Built, in test queue |
| `live` | Approved, active in store |
| `paused` | Temporarily removed from store |
