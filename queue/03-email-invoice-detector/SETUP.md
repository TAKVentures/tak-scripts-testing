# Email Invoice Detector — Setup Guide

**By TAKScripts (TAK Ventures) · takscripts.store**

Automatically scans your Gmail for invoices, receipts, and payment confirmations, then logs them to a clean, branded Google Sheet.

---

## Quick Start (5 minutes)

### 1. Create the Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Rename it to something like **"Invoice Tracker"**.

### 2. Add the Script

1. In your spreadsheet, go to **Extensions → Apps Script**.
2. Delete any code in the editor.
3. Paste the entire contents of `email-invoice-detector.gs`.
4. Click the **Save** icon (or Ctrl+S / Cmd+S).
5. Close the Apps Script editor tab.

### 3. Authorize Permissions

1. Reload your spreadsheet — you should see **🕷 TAKScripts** in the menu bar.
2. Click **🕷 TAKScripts → ⚙️ Settings**.
3. Google will ask you to authorize the script. Click **Continue**.
4. Choose your Google account.
5. If you see "Google hasn't verified this app", click **Advanced → Go to [project name]**.
6. Click **Allow** to grant permissions.

### 4. Configure Settings

Once authorized, the settings sidebar will appear. Configure:

- **Gmail Label** — Leave empty to scan your inbox, or enter a label name to scan only labeled emails.
- **Lookback** — How many days back to scan (default: 7 days).
- **Scan Frequency** — How often the auto-scan runs (default: every hour).
- **Auto-scan toggle** — Enable to create an automatic time-driven trigger.
- **Detection Categories** — Check which types to detect (Invoice, Receipt, Payment Confirmation, etc.).
- **Amount Alerts** — Enable to get email notifications when an invoice exceeds your threshold.
- **Threshold** — Dollar amount that triggers an alert (default: $500).
- **Alert Email** — Where to send high-amount alerts (defaults to your Gmail).

Click **Save Settings** when done.

### 5. Run Your First Scan

- Click **🕷 TAKScripts → 🧪 Test Run** to preview what would be detected (nothing is logged).
- Click **🕷 TAKScripts → ▶️ Run Scan Now** to scan and log results.
- Click **🕷 TAKScripts → 📊 View Invoice Log** to see the results sheet.

---

## Features

### Detection Engine

The script scans email subjects and bodies for keywords in these categories:

| Category | Example Keywords |
|---|---|
| Invoice | "invoice", "inv #", "invoice number" |
| Receipt | "receipt", "payment receipt", "e-receipt" |
| Payment Confirmation | "payment confirmation", "payment received" |
| Order Confirmation | "order confirmation", "order #" |
| Subscription | "subscription", "recurring payment", "renewal" |
| Refund | "refund", "credit issued", "money back" |

### Amount Extraction

Detects currency amounts in these formats:
- `$1,234.56`
- `USD 100.00` / `100.00 USD`
- `EUR 50.00` / `GBP 75.00`
- `£50.00` / `€50.00`

The highest detected amount is logged.

### Idempotency

Every processed email's message ID is stored. Re-running the scan (manually or via trigger) will never create duplicate entries. The script keeps the last 5,000 processed IDs to stay within Apps Script property size limits.

### Alert Emails

When enabled, the script sends a branded HTML email alert whenever an invoice amount exceeds your configured threshold. Alerts include the sender, subject, amount, category, and date.

### Test Run Mode

Run **🧪 Test Run** to see exactly what the scanner would detect without logging anything or sending alerts. Use this to verify your settings before enabling auto-scan.

---

## Sheet Layout

The script automatically creates a **📊 Invoice Log** sheet with these columns:

| Column | Description |
|---|---|
| Date | When the email was received |
| Sender | Email address of the sender |
| Subject | Email subject line |
| Amount | Detected currency amount (display) |
| Amount (Numeric) | Parsed numeric value for sorting/summing |
| Category | Detection category (Invoice, Receipt, etc.) |
| Matched Keyword | Which keyword triggered detection |
| Message ID | Internal Gmail ID (hidden column) |

Rows alternate white/light gray. Categories are color-coded (green for receipts, amber for invoices, red for refunds, blue for orders).

---

## Permissions Required

| Permission | Why |
|---|---|
| Gmail (read, send) | Read inbox to detect invoices; send alert emails |
| Google Sheets | Create and write to the invoice log sheet |
| Script triggers | Create time-driven triggers for auto-scanning |
| Script properties | Store settings and processed message IDs |

---

## Troubleshooting

### "No invoices found" but I know there are some
- Increase the **Lookback** days in settings.
- Make sure the correct **categories** are checked.
- Check that your emails contain the expected keywords in the subject or body.
- If using a Gmail label filter, verify the label name is correct.

### Duplicate entries appearing
This should not happen due to the message ID tracking. If it does:
1. Open **Extensions → Apps Script**.
2. Run the function `resetProcessedIds` from the editor.
3. Delete existing rows from the log sheet.
4. Run the scan again.

### Auto-scan not running
- Open **Extensions → Apps Script → Triggers** (clock icon on the left).
- Verify a trigger for `runScan` exists.
- Check that the trigger is not in a failed state.
- Re-enable auto-scan in settings to recreate the trigger.

### Alert emails not sending
- Confirm **High-amount email alerts** is toggled on.
- Verify the **Alert Email** address is correct.
- Check that the **Threshold** is set to a reasonable value.
- Look at **Extensions → Apps Script → Executions** for error logs.

---

## Advanced: Manual Functions

These functions are available from the Apps Script editor for advanced users:

| Function | Purpose |
|---|---|
| `resetProcessedIds()` | Clears all tracked message IDs so the next scan re-processes everything |
| `removeTrigger_()` | Manually removes all scan triggers |

---

## License & Support

Part of the **TAKScripts** collection by TAK Ventures.

- Store: [takscripts.store](https://takscripts.store)
- Questions? Reach out via the store.
