# TEST — Email Invoice Detector

## Pre-Setup
None. Script scans Gmail directly. The test account (tak.testingdemo@gmail.com) should have
at least one email in the inbox containing words like "invoice", "receipt", or "payment".
If not, send a test email to the account with subject "Invoice #1234 — $50.00" before testing.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Scan last N days | 30 |
| Alert email | tak.testingdemo@gmail.com |
| Amount threshold | 0 (alert on everything) |
| Auto-scan schedule | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → ▶️ Scan Now
7. ✅ `📊 Expense Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ At least one invoice row in row 4+, stats show THIS MONTH or ALL TIME ≥ 1
9. TAKScripts → ▶️ Scan Now a second time
10. ✅ No duplicate rows added (deduplication working), no crash

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- At least one invoice detected and logged
- Second scan adds no duplicates

## FAIL — note exactly what broke and at which step
