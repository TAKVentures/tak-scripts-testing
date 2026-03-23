# TAKScripts Testing Protocol

This document defines the standard testing process for all TAKScripts products.
Follow this for every new script — scripts 1–11 and every script after.

---

## Philosophy

Test from the **customer's perspective**, not the developer's.

Customers never run functions in the Apps Script editor. They:
1. Open a Google Sheet
2. Paste the script via Extensions → Apps Script
3. Use the menu to configure and run it
4. Look at their sheet to see results

A script passes testing when **that flow works end to end**. Edge cases and
developer-facing functions are secondary — ship the happy path, patch the rest
from real customer feedback.

---

## Folder Structure (per script)

```
queue/{slug}/
  script.gs     ← the script to test
  TEST.md       ← test card with settings + steps for this specific script
  SETUP.md      ← customer-facing setup instructions (optional, for reference)

approved/{slug}/
  script.gs     ← copy here after testing passes
```

---

## How to Write a TEST.md for a New Script

Copy `TEST_TEMPLATE.md` from the repo root and fill in:

1. **Pre-Setup** — anything that must exist before the script runs
   (test data in the sheet, a Gmail email to detect, a Drive folder to scan)
2. **Settings** — every field in the Settings sidebar, with a concrete test value
3. **Test Steps** — follow the standard 10-step customer journey below,
   adding script-specific verification steps at 8–11
4. **PASS Criteria** — the specific things that must be true to call it a pass

### Standard 10-Step Customer Journey

Every script uses this base flow. Add script-specific steps between 8 and 10.

```
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste script.gs → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ TAKScripts menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter test values → Save
6. Run [main action] from the TAKScripts menu
7. ✅ 📊 [Dashboard Name] sheet exists with gold stats bar in rows 1-2, headers in row 3
8. ✅ [Script-specific: at least one data row, correct output, etc.]
9. Run [main action] a second time
10. ✅ No crash, no duplicate rows, stats updated
```

---

## Browser Claude Instructions

When running a test session, give Browser Claude this prompt:

```
You are testing a TAKScripts Google Apps Script for quality assurance.

1. Read the TEST.md file in queue/{slug}/
2. Complete every step exactly as written — do not skip or improvise
3. Use the test values provided in the Settings section
4. After completing all steps, report:
   - PASS or FAIL
   - If FAIL: which step failed and what the error/symptom was
   - Any unexpected behavior worth noting, even if it didn't cause a failure

Keep the session focused. Do not explore functions beyond what TEST.md specifies.
```

---

## After Testing

| Result | Action |
|---|---|
| ✅ PASS | Copy `script.gs` to `approved/{slug}/script.gs` → commit → push → upload to Supabase Storage → set `is_active = true` |
| ❌ FAIL | File issue with Claude Code (include step # and error) → fix applied to source → re-copy to queue → re-commit → re-test |
| ⚠️ MINOR | Fix if customer-facing, ship if editor-only (e.g. getUi() in editor context) |

---

## What a TEST.md Does NOT Cover

These are known gaps — accepted for v1, fix post-launch from real feedback:

- First-run OAuth consent screen flow
- Trigger/auto-schedule functionality (daily/weekly timers)
- Bad input / error handling (wrong folder ID, invalid email, etc.)
- Settings persistence across sidebar close/reopen
- Empty state (no matching Gmail threads, no inventory data, etc.)
