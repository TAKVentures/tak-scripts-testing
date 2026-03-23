# TEST — Bulk Email Unsubscriber

## Pre-Setup
The script scans Gmail for subscription/newsletter emails and surfaces unsubscribe options.
The test account should have promotional or newsletter emails in Gmail.
No pre-setup needed — the script finds them automatically.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Scan last N days | 90 |
| Min emails from sender | 2 |
| Include Promotions tab | ON |
| Include Updates tab | ON |
| Include Social tab | OFF |
| Excluded senders | (leave blank) |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → ▶️ Scan Subscriptions
7. ✅ `📊 Subscription Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ Subscription rows appear in row 4+ with sender, frequency, and unsubscribe status
9. ✅ Stats show FOUND ≥ 1
10. TAKScripts → ▶️ Scan Subscriptions a second time
11. ✅ Dashboard refreshes cleanly (rows rewritten, not doubled), stats update, no crash

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- At least one subscription detected
- Second scan rewrites rows cleanly (not appended twice)

## FAIL — note exactly what broke and at which step
