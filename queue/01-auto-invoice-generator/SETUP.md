# Auto Invoice Generator — Setup Guide

**By TAKScripts (TAK Ventures)**
Version 1.0

---

## What You Get

A complete invoicing system inside Google Sheets:

- Generate branded PDF invoices from client data in your spreadsheet
- Auto-email invoices to clients with a professional HTML email
- Track every invoice (Draft / Sent / Paid) in a log sheet
- Configurable tax rate, payment terms, invoice numbering, and email template
- All PDFs saved to an organized Google Drive folder
- Test Run mode to preview invoices without sending emails

---

## Installation (5 minutes)

### Step 1: Create the Spreadsheet

1. Open [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet
2. Name it something like "Invoice Manager"

### Step 2: Open the Script Editor

1. In your spreadsheet, go to **Extensions > Apps Script**
2. Delete any code already in the editor
3. Paste the entire contents of `auto-invoice-generator.gs` into the editor
4. Click the **Save** icon (or Ctrl/Cmd + S)
5. Name the project "Auto Invoice Generator"

### Step 3: Authorize the Script

1. Close the Apps Script editor and refresh your spreadsheet
2. You should see a **🕷 TAKScripts** menu appear in the menu bar
3. Click **🕷 TAKScripts > ⚙️ Settings**
4. Google will ask you to authorize the script — click through the prompts:
   - "This app isn't verified" — click **Advanced > Go to Auto Invoice Generator (unsafe)**
   - Review the permissions and click **Allow**

> **Why these permissions?** The script needs access to Google Sheets (read/write client data), Google Docs (create invoice documents), Google Drive (save PDFs), and Gmail (send invoice emails). It only accesses files it creates.

### Step 4: Configure Your Business

1. Open **🕷 TAKScripts > ⚙️ Settings**
2. Fill in your business details:
   - **Business Name** — appears on invoices and emails
   - **Business Address** — appears on invoice header
   - **Business Email** — appears on invoice header
   - **Logo URL** (optional) — public URL to your logo image
3. Set your invoice preferences:
   - **Invoice Prefix** — e.g., "INV-" produces INV-1001, INV-1002, etc.
   - **Next Number** — starting invoice number
   - **Tax Rate** — percentage (0 for no tax)
   - **Due Days** — number of days until payment is due (default: 30)
   - **Payment Terms** — text shown on the invoice
4. Configure email:
   - **Auto-send** toggle — when ON, invoices are emailed automatically
   - **Email Subject** — supports template variables
   - **Email Body** — supports template variables
5. Click **Save Settings**

### Step 5: Initialize Sheets

1. In the Settings sidebar, click **Initialize Sheets** at the bottom
2. This creates two pre-formatted sheets:
   - **👥 Clients** — where you add client data
   - **📊 Invoice Log** — where invoices are tracked

---

## Usage

### Adding Clients

Open the **👥 Clients** sheet and fill in one row per client:

| Column | Description | Example |
|--------|-------------|---------|
| Client Name | Company or individual name | Acme Corp |
| Email | Where to send the invoice | billing@acme.com |
| Address | Client billing address | 123 Main St, Springfield, IL |
| Service Description | What you're billing for | Website Development |
| Rate ($) | Per-unit or per-hour rate | 150 |
| Hours/Qty | Number of hours or units | 10 |
| Notes | Additional notes for the invoice | Monthly retainer |

### Generating an Invoice

1. Go to **🕷 TAKScripts > ▶️ Generate Invoice**
2. Enter the row number of the client you want to invoice
3. Or type "all" to generate invoices for every client
4. The script will:
   - Create a branded Google Doc invoice
   - Convert it to PDF
   - Save the PDF to a **📁 TAKScripts — Invoices** folder in your Drive
   - Email the PDF to the client (if auto-email is enabled)
   - Log the invoice in the **📊 Invoice Log** sheet

### Test Run

Go to **🕷 TAKScripts > 🧪 Test Run** to generate a sample invoice without sending any emails. The PDF is still saved to Drive so you can review the formatting.

### Marking Invoices as Paid

Two ways to mark an invoice as paid:

1. **From the menu:** Select a row in the Invoice Log, then the status will update
2. **Manual edit:** Type "Paid" directly in the Status column — the styling will auto-apply

### Viewing the Invoice Log

Go to **🕷 TAKScripts > 📊 View Invoice Log** to jump to the log sheet. Each row shows:

- Invoice number
- Client name and email
- Total amount
- Date created and due date
- Status (color-coded: Draft, Sent, Paid, Error)
- Clickable link to the PDF in Drive

---

## Email Template Variables

Use these in the email subject and body:

| Variable | Replaced With |
|----------|--------------|
| `{{client_name}}` | Client's name from the sheet |
| `{{invoice_number}}` | Full invoice number (e.g., INV-1001) |
| `{{total}}` | Invoice total with dollar sign |
| `{{due_date}}` | Payment due date |
| `{{business_name}}` | Your business name from settings |

---

## File Structure

After generating invoices, your Google Drive will have:

```
My Drive/
  📁 TAKScripts — Invoices/
    INV-1001.pdf
    INV-1002.pdf
    ...
```

---

## Troubleshooting

### Menu doesn't appear
Refresh the spreadsheet. The menu loads on open. If it still doesn't appear, open **Extensions > Apps Script** and run the `onOpen` function manually.

### "Authorization required" error
You need to re-authorize. Go to **Extensions > Apps Script**, click Run on any function, and follow the authorization prompts.

### Emails not sending
- Make sure the **Auto-send** toggle is ON in Settings
- Check that the client has a valid email address
- Gmail has a daily sending limit (100 for free accounts, 1500 for Workspace). Check your quota at [Google Apps Script Dashboard](https://script.google.com).

### Invoice PDF looks wrong
- Run a Test Run first to preview
- Check that your business details are filled in Settings
- The Google Doc is created temporarily and then deleted after PDF conversion. The PDF is the permanent copy.

### Tax not showing on invoice
- Set a tax rate greater than 0 in Settings
- Tax is calculated as: (Rate x Hours/Qty) x (Tax Rate / 100)

---

## Permissions Overview

| Permission | Why |
|-----------|-----|
| Google Sheets | Read client data, write invoice log |
| Google Docs | Create invoice documents (temporary) |
| Google Drive | Save PDF invoices to a folder |
| Gmail | Send invoice emails with PDF attachment |
| Script Properties | Store your settings between sessions |

---

## Support

Questions or issues? Visit [takscripts.store](https://takscripts.store) for help.

---

*Powered by TAKScripts · takscripts.store*
