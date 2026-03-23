# TEST — Inventory Low-Stock Alert

## Pre-Setup
After step 3 (reload), a `📊 Inventory Dashboard` sheet is created automatically.
The script expects inventory data starting at row 4. Add these test rows manually after the sheet appears:

| Item Name | SKU | Category | Reorder Point | Current Stock | Supplier | Supplier Email |
|---|---|---|---|---|---|---|
| Widget A | SKU-001 | Electronics | 10 | 25 | Acme Supply | supplier@example.com |
| Widget B | SKU-002 | Electronics | 10 | 5 | Acme Supply | supplier@example.com |
| Widget C | SKU-003 | Hardware | 20 | 0 | Bob's Parts | bob@example.com |

Widget B (stock 5, reorder 10) and Widget C (stock 0) should trigger alerts.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Alert email | tak.testingdemo@gmail.com |
| Check frequency | Manual only |
| Supplier emails | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → ▶️ Initial Setup (if available) — otherwise skip
7. Add the three test inventory rows to `📊 Inventory Dashboard` starting at row 4
8. TAKScripts → ▶️ Check Stock Now
9. ✅ Dashboard stats update — LOW STOCK and OUT OF STOCK counts are non-zero
10. ✅ Widget B row highlighted amber, Widget C row highlighted red
11. ✅ `📋 Alert History` sheet has entries for this run
12. TAKScripts → ▶️ Check Stock Now a second time
13. ✅ No crash, no duplicate alert history rows

## PASS Criteria
- Menu appears after reload
- Stats bar shows correct low/out stock counts
- Row color-coding works (amber = low, red = out)
- Alert history logged
- Second run doesn't duplicate

## FAIL — note exactly what broke and at which step
