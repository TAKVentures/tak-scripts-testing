# TEST — Smart OOO Auto-Reply

## Pre-Setup
The script sends auto-replies to incoming emails during a set date range.
For testing, set the date range to include today so the OOO is "active".
The test account should have at least one recent unread email from a real address.
If not, send a test email to tak.testingdemo@gmail.com before testing.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Start Date | today's date (e.g. 2026-03-23) |
| End Date | one week from today (e.g. 2026-03-30) |
| Reply Subject | Out of Office — Back Soon |
| Reply Message | Hi, I'm currently out of the office. I'll respond when I return. — TAK Test |
| Skip domains | noreply, no-reply, mailer-daemon, newsletter, notifications |
| Max replies per sender | 1 |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → ▶️ Check & Reply
7. ✅ `📊 OOO Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ If any qualifying emails found: row logged, TOTAL REPLIES ≥ 1
9. ✅ If no qualifying emails: dashboard created with stats showing `—`, no crash
10. TAKScripts → ▶️ Check & Reply a second time
11. ✅ No duplicate replies sent to same sender (dedup working), no crash

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- No crash whether or not qualifying emails exist
- Second run does not re-reply to same senders

## FAIL — note exactly what broke and at which step
