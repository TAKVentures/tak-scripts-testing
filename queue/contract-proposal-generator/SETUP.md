# Contract & Proposal Generator — Setup Guide

**by TAKScripts (TAK Ventures)**

Generate branded proposals and contracts from a spreadsheet, convert to PDF, and email to clients — all without leaving Google Sheets.

---

## Quick Start (5 minutes)

### 1. Create a Google Sheet

- Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
- Name it something like **"Proposals & Contracts"**

### 2. Add the Script

- Go to **Extensions > Apps Script**
- Delete any existing code in the editor
- Copy and paste the entire contents of `contract-proposal-generator.gs`
- Click **Save** (disk icon or Ctrl+S)
- Close the Apps Script editor tab

### 3. Reload the Sheet

- Refresh the spreadsheet page (F5 or Ctrl+R)
- You should see the **🕷 TAKScripts** menu appear in the menu bar
- If prompted, click **Continue** and authorize the script

### 4. Run Initial Setup

- Click **🕷 TAKScripts > ⚙️ Settings**
- In the sidebar, scroll down and click **🔧 Run Initial Setup**
- This creates the **👥 Clients** and **📋 Document Log** sheets with proper formatting

### 5. Configure Your Business

In the Settings sidebar, fill in:

| Field | What to enter |
|-------|--------------|
| Business Name | Your company name (appears on documents) |
| Address | Full business address |
| Email | Business contact email |
| Phone | Business phone number |
| Logo URL | *(Optional)* Direct link to your logo image |
| Proposal Prefix | Default: `PROP-` — change to match your numbering |
| Contract Prefix | Default: `CONTRACT-` — change to match your numbering |
| Terms & Conditions | Your standard T&C text |
| Payment Terms | Your default payment terms |
| Email Subject | Template for outgoing emails |
| Email Body | Template for outgoing emails |

Click **Save Settings** when done.

---

## How to Use

### Adding Clients

1. Go to the **👥 Clients** sheet
2. Fill in a row with the client's information:
   - **Company** — Client company name
   - **Contact Name** — Person you're sending to
   - **Email** — Where to send the document
   - **Address** — Client address (appears on document)
   - **Project Description** — Brief project summary
   - **Scope** — Detailed scope of work
   - **Deliverables** — Comma-separated list of deliverables
   - **Timeline** — Project timeline
   - **Rate/Price** — Total amount or rate
   - **Payment Terms** — Payment terms for this project (or leave blank to use defaults)

### Generating a Proposal

1. Go to the **👥 Clients** sheet
2. **Click on the row** of the client you want
3. Click **🕷 TAKScripts > 📝 Generate Proposal**
4. Confirm the details in the popup
5. The script will:
   - Create a branded Google Doc
   - Convert it to PDF
   - Save the PDF to your **📁 TAKScripts — Proposals** folder in Drive
   - Email the PDF to the client
   - Log everything in the **📋 Document Log** sheet

### Generating a Contract

Same process, but choose **📄 Generate Contract** from the menu. The document formatting and numbering will reflect "Contract" instead of "Proposal."

### Test Run

- Click **🕷 TAKScripts > 🧪 Test Run**
- Generates the document and PDF **without sending an email**
- Great for previewing before sending to a real client

---

## Managing Documents

### Document Log

The **📋 Document Log** sheet tracks every generated document:

| Column | Description |
|--------|-------------|
| Doc Number | Auto-generated (PROP-001, CONTRACT-001, etc.) |
| Client | Company name |
| Type | Proposal or Contract |
| Date Created | When it was generated |
| Amount | Price from the client row |
| Status | Draft, Sent, Signed, or Expired |
| PDF Link | Direct link to the PDF in Drive |

### Updating Status

To mark a document as **Signed** or **Expired**:

1. Go to the **📋 Document Log** sheet
2. Click on the row you want to update
3. Open the Settings sidebar (**⚙️ Settings**)
4. Click **✅ Mark Selected Doc as Signed** or **❌ Mark Selected Doc as Expired**

Status colors:
- **Sent** — Green
- **Draft** — Amber
- **Signed** — Blue
- **Expired** — Red

---

## Email Template Variables

Use these in the Email Subject and Email Body fields:

| Variable | Replaced with |
|----------|--------------|
| `{{contactName}}` | Client's contact name |
| `{{type}}` | "proposal" or "contract" |
| `{{business}}` | Your business name |
| `{{docNumber}}` | The document number (e.g., PROP-001) |

---

## Where Documents Are Saved

All PDFs are saved to a Google Drive folder called **📁 TAKScripts — Proposals**. The folder is created automatically the first time you generate a document. You can find it in your Drive root.

---

## Authorization

The first time you run the script, Google will ask you to authorize it. The script needs access to:

- **Google Sheets** — To read client data and write the document log
- **Google Docs** — To create the document
- **Google Drive** — To save the PDF
- **Gmail** — To send the email with the PDF attached

Click **Advanced > Go to (your project)** if you see a warning screen, then click **Allow**.

---

## Troubleshooting

| Issue | Solution |
|-------|---------|
| Menu doesn't appear | Refresh the sheet. If still missing, re-open the Apps Script editor and run `onOpen` manually. |
| "Business name not set" error | Open Settings and fill in your business name. |
| "Select a client first" error | Make sure you're on the **👥 Clients** sheet and have clicked on a data row. |
| Email not received | Check spam. Make sure the client email address is correct in the sheet. |
| PDF not in Drive | Look for the **📁 TAKScripts — Proposals** folder in your Drive root. |
| Document numbers restarted | Script Properties may have been cleared. Numbers will continue from where they left off in normal use. |

---

## Support

Questions? Visit [takscripts.store](https://takscripts.store) or email support.

---

*Contract & Proposal Generator v1.0 — by TAK Ventures*
