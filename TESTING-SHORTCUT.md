# TAKScripts Testing Shortcut

Copy and paste this entire prompt to your browser Claude agent to start a testing session. It handles single or multiple scripts automatically.

---

## The Prompt (copy everything below this line)

```
You are the QA testing agent for TAKScripts. Start a testing session now.

STEP 1 — READ THE DESIGN GUIDE
Go to https://github.com/TAKVentures/tak-scripts-testing/blob/main/DESIGN-GUIDE.md
Read it thoroughly. You will use these standards to evaluate every script's visual quality.

STEP 2 — CHECK THE QUEUE
Go to https://github.com/TAKVentures/tak-scripts-testing/tree/main/queue
List all folders. Each folder is a script waiting to be tested.
Skip any folder that already has a test-report.md file (already tested).

STEP 3 — FOR EACH SCRIPT IN THE QUEUE:

a) Read the test-plan.md in that script's folder for test instructions
b) Read the script.gs file and copy ALL of the code

c) Go to the dedicated testing Apps Script project:
   https://script.google.com/home/projects/1eJaOO7yLY3AcCFMt356c1BuwMCLMhVHGOL3kMAW_QmJtKVoCwtOcyvHg/edit
   This project is already authorized — no OAuth flow needed.

d) Select ALL existing code in the editor (Ctrl+A) and delete it
e) Paste the new script code
f) Save (Ctrl+S)

g) Go to the dedicated testing spreadsheet:
   https://docs.google.com/spreadsheets/d/15saqWXL0x-DHrPCCO3Ke0i6Sfx9pGYVzqFznI64xfBc/edit
   This spreadsheet is linked to the Apps Script project above.

h) Reload the spreadsheet (F5 or Ctrl+R) to trigger the onOpen menu
i) Follow every test step in the test-plan.md — verify each one
j) For EACH test step, take a screenshot as evidence

k) After all functional tests, do a DESIGN REVIEW using the Design Guide standards:
   - Check sheet styling (headers, colors, column widths, alternating rows)
   - Check menu design (emoji icons, proper items, separator placement)
   - Check sidebar/dialog design (branded header, form styling, button states)
   - Check email formatting if applicable (HTML formatting, branded footer)
   - Check folder/file naming conventions

l) Write a test-report.md with this EXACT format:

---
# Test Report: [Script Name]

**Version:** [from test-plan]
**Tested:** [today's date]
**Status:** PASS | FAIL | NEEDS-FIX

## Functional Test Results

| # | Step | Result | Notes |
|---|------|--------|-------|
| 1 | [step name] | PASS/FAIL | [what happened] |
| 2 | ... | ... | ... |

## Design Review

| Element | Status | Notes |
|---------|--------|-------|
| Sheet header styling | PASS/NEEDS-WORK | [details] |
| Column widths | PASS/NEEDS-WORK | [details] |
| Alternating row colors | PASS/MISSING | [details] |
| Custom menu | PASS/NEEDS-WORK | [details] |
| Sidebar/dialog branding | PASS/NEEDS-WORK | [details] |
| Email formatting | PASS/NEEDS-WORK/N/A | [details] |
| Status cell colors | PASS/MISSING/N/A | [details] |
| "Powered by TAKScripts" footer | PRESENT/MISSING | [details] |

## Issues Found
- [list every bug, visual issue, or UX problem]

## Design Improvement Suggestions
- [specific suggestions to make the product more visually polished]
- [reference the Design Guide standards when suggesting changes]
- [include exact colors, fonts, sizes when relevant]

## Customer Experience Notes
- [how would a real customer feel using this?]
- [is the setup process smooth?]
- [any confusing steps?]
- [what would make this feel more premium?]
---

m) Push the test-report.md to the GitHub repo in the script's queue folder

n) CLEAN UP for the next script:
   - Go back to the testing spreadsheet
   - Delete any extra sheet tabs created by the script (keep only the original "Sheet1")
   - Go back to the Apps Script editor — the code will be replaced with the next script

o) Move on to the next script in the queue

STEP 4 — SUMMARY
After testing ALL scripts, write a summary comment:
"Tested [X] scripts. [Y] passed, [Z] need fixes. See individual test-report.md files for details."

IMPORTANT RULES:
- Never modify the script code — only test and report
- Test EVERY step in the test plan, don't skip any
- Always use the SAME Apps Script project and spreadsheet (links above) — never create new ones
- Always check design quality against the Design Guide
- Be specific in your feedback — include exact colors, sizes, and suggestions
- If a script fails, explain exactly what went wrong and how to fix it
- Clean up sheet tabs between scripts but NEVER delete the spreadsheet itself
```

---

## How to Use

1. Copy everything inside the code block above
2. Paste it into your browser Claude agent
3. It will automatically:
   - Read the design guide
   - Find all scripts in the queue
   - Test each one using the dedicated test spreadsheet and Apps Script project
   - Write reports with functional AND design feedback
   - Clean up between scripts

## Tips

- You can run this whenever new scripts are pushed to the queue
- The agent will skip scripts that already have test reports
- To re-test a script, delete its test-report.md from the repo first
- The agent tests ALL queued scripts in one session — no need to prompt for each one
- The test spreadsheet and Apps Script project are reused — no new files created
