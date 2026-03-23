# TEST — Follow-Up Nudger

## Pre-Setup
The script scans Gmail Sent folder for emails that haven't received a reply.
The test account should have at least one sent email older than 3 days with no reply.
If not, send an email to a dummy address (e.g. test-noreply@example.com) from the account
a few days before testing, or set `daysBeforeFollowUp` to 0 in settings.

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| Days before follow-up | 0 (catches everything immediately for testing) |
| Max follow-ups per thread | 2 |
| Auto-send follow-ups | OFF |
| Send digest email | OFF |
| Excluded domains | noreply, no-reply, mailer-daemon, notifications |

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. TAKScripts → ▶️ Run Follow-Up Check
7. ✅ `📊 Follow-Up Dashboard` sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ Pending follow-up rows appear in row 4+, stats show PENDING ≥ 1
9. TAKScripts → ▶️ Run Follow-Up Check a second time
10. ✅ No crash, same rows (not duplicated), stats consistent

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with stats bar
- At least one pending follow-up detected
- Second run does not duplicate rows

## FAIL — note exactly what broke and at which step
