# TEST — [Script Name]

## Pre-Setup
<!--
Describe anything that must exist BEFORE the script runs:
- Test data rows to add to a sheet
- A Gmail email the script should detect
- A Google Drive folder to scan
- Any other account state needed

If nothing is needed, write: "None. Script works on a blank sheet."
-->

## Settings (TAKScripts → ⚙️ Settings)
| Field | Value |
|---|---|
| [Setting name] | [Test value] |
| [Setting name] | [Test value] |

<!--
List every field in the Settings sidebar with a concrete test value.
Use tak.testingdemo@gmail.com for any email fields.
Use OFF/disabled for any auto-schedule or auto-send fields.
-->

## Test Steps
1. Open a blank Google Sheet
2. Extensions → Apps Script → paste `script.gs` → Save (Ctrl+S)
3. Reload the Sheet tab
4. ✅ `🕷 TAKScripts` menu appears in the menu bar
5. TAKScripts → ⚙️ Settings → enter values above → Save
6. [Add pre-run steps here if needed, e.g. adding test data rows]
7. TAKScripts → [▶️ Main Action Name]
8. ✅ `📊 [Dashboard Sheet Name]` sheet exists with gold stats bar in rows 1-2, headers in row 3
9. ✅ [Script-specific check: e.g. "at least one data row in row 4+, stats show X ≥ 1"]
10. TAKScripts → [▶️ Main Action Name] a second time
11. ✅ No crash, no duplicate rows, stats updated correctly

## PASS Criteria
- Menu appears after reload
- Dashboard sheet created with gold stats bar
- [Script-specific: describe what correct output looks like]
- Second run produces no duplicates or errors

## FAIL — note exactly what broke and at which step
