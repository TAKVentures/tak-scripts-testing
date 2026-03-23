# Bulk Email Unsubscriber — Setup Guide

**By TAKScripts · TAK Ventures**
**Version 1.0**

---

## What It Does

Scans your Gmail for newsletters and subscription emails, lists them in a clean dashboard inside Google Sheets, and lets you bulk-unsubscribe with one click. It never deletes emails — it only unsubscribes.

---

## Quick Setup (5 minutes)

### 1. Create a Google Sheet

- Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet
- Name it something like **"Email Unsubscriber"**

### 2. Open the Script Editor

- In your new sheet, go to **Extensions → Apps Script**
- Delete any existing code in the editor

### 3. Paste the Script

- Open `bulk-email-unsubscriber.gs` from this folder
- Copy the entire contents
- Paste it into the Apps Script editor, replacing everything

### 4. Save and Authorize

- Click the **Save** icon (or Ctrl/Cmd + S)
- Name the project **"Bulk Email Unsubscriber"**
- Close the Apps Script editor tab and **refresh your spreadsheet**
- You should see the **🕷 TAKScripts** menu appear in the menu bar
- Click any menu item — Google will ask you to authorize the script
- Click **Advanced → Go to Bulk Email Unsubscriber (unsafe)** → **Allow**

> **Note:** The "unsafe" warning is normal for all custom scripts. The script only reads your emails and sends unsubscribe requests — it never deletes anything.

### 5. Configure Settings

- Click **🕷 TAKScripts → ⚙️ Settings**
- Set your preferences:
  - **Scan Period:** How many days back to search (default: 90)
  - **Minimum Emails:** Only show senders with at least this many emails (default: 3)
  - **Categories:** Which Gmail categories to include (Promotions, Updates, Social)
  - **Exclude Senders:** Comma-separated list of senders to never flag
- Click **Save Settings**

---

## How to Use

### Scan for Subscriptions

1. Click **🕷 TAKScripts → ▶️ Scan Subscriptions**
2. Wait for the scan to complete (takes 30–60 seconds depending on inbox size)
3. Review the results on the **📬 Subscription Scanner** tab

### Test Run (Preview Only)

1. Click **🕷 TAKScripts → 🧪 Test Run**
2. This scans your inbox and shows results, but does NOT process any unsubscribes
3. Use this to preview what the script will find before committing

### Unsubscribe from Senders

1. On the **📬 Subscription Scanner** sheet, find the senders you want to unsubscribe from
2. Change the **Status** column dropdown from "Keep" to **"Unsubscribe"**
3. Click **🕷 TAKScripts → ▶️ Scan Subscriptions** again
4. The script will process all rows marked "Unsubscribe"
5. Status will update to **"Done"** (success) or **"Error"** (failed)

### View the Log

- Click **🕷 TAKScripts → 📊 View Log**
- The **📋 Unsubscribe Log** tab shows every action taken, with timestamps and results

---

## Column Reference

### 📬 Subscription Scanner

| Column | Description |
|--------|-------------|
| Sender | Display name of the sender |
| Email Address | The sender's email address |
| Emails Found | Number of emails received from this sender |
| Frequency | How often they email (Daily, Weekly, Monthly, etc.) |
| Last Received | Date of the most recent email |
| Status | Keep, Unsubscribe, Done, or Error (dropdown) |
| Unsubscribe Link | The detected unsubscribe URL or mailto link |

### 📋 Unsubscribe Log

| Column | Description |
|--------|-------------|
| Timestamp | When the action was performed |
| Sender | Display name |
| Email Address | Sender's email |
| Method | URL or Email |
| Result | Success or Failed |
| Details | HTTP response code or error message |

---

## How Detection Works

The script finds subscriptions by looking for:

1. **List-Unsubscribe headers** — The standard email header that indicates a mailing list
2. **Gmail Promotions category** — Emails Gmail has classified as promotional
3. **Gmail Updates category** — Emails Gmail has classified as updates
4. **Gmail Social category** — Social notification emails (optional)

It then aggregates by sender and only shows senders who have emailed you at least the minimum threshold number of times.

---

## How Unsubscribing Works

The script supports two unsubscribe methods:

- **URL unsubscribe:** Visits the unsubscribe link via HTTP GET (the standard one-click unsubscribe method)
- **Mailto unsubscribe:** Sends an email to the unsubscribe address with the appropriate subject line

Some subscriptions may not have a detectable unsubscribe link — these will show "(none found)" in the link column. You can manually unsubscribe from those by opening the emails directly.

---

## Permissions Required

| Permission | Why |
|-----------|-----|
| Gmail (read & send) | Read emails to find subscriptions; send unsubscribe emails for mailto links |
| Sheets (read & write) | Create and update the scanner and log sheets |
| External URLs | Fetch unsubscribe links to process URL-based unsubscribes |
| Script Properties | Store your settings between sessions |

---

## Troubleshooting

**Menu doesn't appear:**
Refresh the spreadsheet. The menu loads when the sheet opens.

**Authorization error:**
Click Advanced → Go to [project name] (unsafe) → Allow. This is required for all custom Apps Scripts.

**"Error" status on some rows:**
Some unsubscribe links may be expired, require CAPTCHA, or need you to confirm on a webpage. Open the link manually in those cases.

**Scan finds nothing:**
Try increasing the scan period in Settings, lowering the minimum email count, or enabling more categories.

**Script times out:**
Google Apps Script has a 6-minute execution limit. If you have a very large inbox, try reducing the scan period to 30 days.

---

## Safety Notes

- The script **never deletes emails** — it only processes unsubscribe links
- The script **never modifies existing emails** — it only reads them
- Unsubscribe emails are sent from your Gmail account so the sender can identify you
- All actions are logged on the **📋 Unsubscribe Log** sheet for your records
- You can always re-subscribe to any sender by visiting their website

---

*Powered by TAKScripts · [takscripts.store](https://takscripts.store)*
