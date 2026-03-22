# TAK Scripts — Testing Pipeline

This repo is the QA handoff layer between **Claude Code** (builds scripts) and **Browser Claude** (tests scripts in a live Google account).

## How It Works

```
Claude Code builds script → pushes to /queue/{slug}/
Browser Claude picks it up → tests in Apps Script → writes test-report.md
Claude Code reads report → fixes issues → pushes update or moves to /approved/
```

## Directory Structure

```
queue/                        ← Scripts waiting to be tested
  {slug}/
    script.gs                 ← The Google Apps Script code
    test-plan.md              ← What to test and how to verify
    test-report.md            ← Browser Claude writes results here

approved/                     ← Scripts that passed all tests
  {slug}/
    script.gs                 ← Final approved code
    test-report.md            ← Passing test report
```

## For Browser Claude

### Your Job
1. Check `/queue/` for new scripts (folders without a `test-report.md`)
2. For each script:
   - Read `test-plan.md` for instructions
   - Read `script.gs` for the code
   - Create a new Apps Script project at `script.google.com`
   - Paste the code, save, and authorize
   - Run each test step from the test plan
   - Verify results across Google tools (Sheets, Gmail, Drive, etc.)
   - Record a GIF if possible
3. Write your results to `test-report.md` in the same folder
4. Use this format for the report:

```markdown
# Test Report: {Script Name}

**Version:** {version}
**Tested:** {date}
**Status:** PASS | FAIL | NEEDS-FIX

## Results

| Step | Result | Notes |
|------|--------|-------|
| Script runs without errors | PASS | Execution completed in 2.3s |
| Custom menu appears | PASS | "TAKScripts" menu visible after reload |
| ... | ... | ... |

## Issues Found
- (list any bugs or improvements needed)

## Recommendations
- (suggestions for the developer)
```

### How to Push Your Report
Use the GitHub API or navigate to the repo in your browser and create/edit the `test-report.md` file directly on GitHub.

## For Claude Code

### Pushing a Script for Testing
1. Create `/queue/{slug}/script.gs` with the full script code
2. Create `/queue/{slug}/test-plan.md` with test steps
3. Commit and push
4. Wait for Browser Claude to write `test-report.md`

### After Testing Passes
1. Read the test report
2. If PASS: move the folder to `/approved/{slug}/`
3. Push to the private `TAK-Scripts` repo
4. Upload to Supabase Storage
5. Set `is_active = true` in the products table

### After Testing Fails
1. Read the test report and issues
2. Fix the script locally
3. Push the updated `script.gs` to `/queue/{slug}/`
4. Delete the old `test-report.md` so Browser Claude re-tests

---

*Part of the TAK Scripts QA pipeline by TAK Ventures*
