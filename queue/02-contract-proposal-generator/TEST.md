# TEST — Contract & Proposal Generator

## Pre-Setup
After step 3 (reload), a `👥 Clients` sheet will be created automatically. Add one test client row:

| Client Name | Contact Name | Email | Address | City | State | ZIP |
|---|---|---|---|---|---|---|
| Acme Corp | Jane Smith | jane@acme.com | 456 Oak Ave | Austin | TX | 78702 |

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Business Name | TAK Test Co |
| Business Email | tak.testingdemo@gmail.com |
| Business Address | 123 Spider St, Austin TX 78701 |
| Business Phone | (512) 555-0100 |
| Proposal Prefix | PROP- |
| Contract Prefix | CONTRACT- |
| Default Terms | Payment due within 30 days |
| Auto-email documents | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. Add the test client row to the `👥 Clients` sheet
7. TAKScripts → 📝 Generate Proposal → select Acme Corp → fill in one service line (e.g. "Web Design", $1500) → Generate
8. ✅ `📊 Pipeline Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
9. ✅ Proposal row appears in row 4+, stats show PROPOSALS ≥ 1
10. TAKScripts → 📄 Generate Contract → select Acme Corp → Generate
11. ✅ Contract row appended, stats updated, no crash

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- Proposal and contract rows logged correctly
- Stats update after each generation

## FAIL — note exactly what broke and at which step
