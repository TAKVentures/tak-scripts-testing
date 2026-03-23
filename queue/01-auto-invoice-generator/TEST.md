# TEST — Auto Invoice Generator

## Pre-Setup
The script needs a Clients sheet with at least one row before generating an invoice.
After step 3 (reload), a `👥 Clients` sheet will be created automatically — add one row of test data:

| Client Name | Email | Address | City | State | ZIP |
|---|---|---|---|---|---|
| Test Client | testclient@example.com | 123 Main St | Austin | TX | 78701 |

## Settings (Extensions → Apps Script → run `showSettings` OR TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Business Name | TAK Test Co |
| Business Email | tak.testingdemo@gmail.com |
| Business Address | 123 Spider St, Austin TX 78701 |
| Tax Rate | 8 |
| Due Days | 30 |
| Invoice Prefix | INV- |
| Starting Invoice # | 1001 |
| Auto-email invoices | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. Add the test client row to the `👥 Clients` sheet
7. TAKScripts → ▶️ Generate Invoice → select the test client → fill in one line item (e.g. "Consulting", qty 1, price 500) → Generate
8. ✅ `📊 Invoice Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
9. ✅ Invoice row appears in row 4+, stats show TOTAL ≥ 1
10. Generate a second invoice for the same client
11. ✅ New row appended, no crash, stats updated

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- Invoice rows logged correctly
- Stats update on second run

## FAIL — note exactly what broke and at which step
