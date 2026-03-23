# TEST — VIP Priority Alert

## Pre-Setup
After step 3 (reload), a `⭐ VIP Contacts` sheet is created. Add at least one VIP contact
whose email address has actually sent an email to the test account recently:

| Name | Email | Priority |
|---|---|---|
| Test VIP | [an email address that has emailed the account] | Critical |

If no real VIP emails exist, add the test account's own address (tak.testingdemo@gmail.com)
and send a test email to itself with subject containing "urgent" before testing.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Alert email | tak.testingdemo@gmail.com |
| Keywords | urgent, asap, important, action required |
| Check interval | 5 minutes |
| Quiet hours | OFF |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. Add the VIP contact row to the `⭐ VIP Contacts` sheet
7. TAKScripts → ▶️ Check Now
8. ✅ `📊 VIP Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
9. ✅ Stats show VIPS TRACKED ≥ 1
10. If a VIP email was found: ✅ row logged in dashboard, TOTAL ALERTS ≥ 1
11. TAKScripts → ▶️ Check Now a second time
12. ✅ No crash, no duplicate rows for same emails

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- VIPS TRACKED stat reflects the VIP Contacts list
- No crash on repeated runs

## FAIL — note exactly what broke and at which step
