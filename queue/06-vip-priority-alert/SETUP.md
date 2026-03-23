# VIP Priority Alert — Setup Guide

**By TAKScripts (TAK Ventures)**
Version 1.0

---

## What It Does

VIP Priority Alert monitors your Gmail inbox and instantly notifies you when emails arrive from your most important contacts. Never miss a message from your boss, top client, or key partner again.

**Features at a glance:**
- Track VIP contacts with three priority levels (Critical, Important, Normal)
- Get branded HTML email alerts with message preview
- Auto-star and label VIP emails in Gmail
- Keyword triggers — flag emails by subject regardless of sender
- Quiet hours to prevent late-night notifications
- Full alert log with color-coded priority tracking
- Configurable check frequency (1–15 minutes)

---

## Installation (5 minutes)

### Step 1: Create the Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it something like **"VIP Priority Alert"**

### Step 2: Add the Script

1. In your spreadsheet, go to **Extensions → Apps Script**
2. Delete any existing code in the editor
3. Copy the entire contents of `vip-priority-alert.gs` and paste it in
4. Click **Save** (Ctrl+S / Cmd+S)
5. Name the project **"VIP Priority Alert"**

### Step 3: Authorize Permissions

1. Close the Apps Script editor and **reload your spreadsheet**
2. You will see a **🕷 TAKScripts** menu appear in the menu bar
3. Click **🕷 TAKScripts → ▶️ Start Monitoring**
4. Google will ask you to authorize the script — click through and allow:
   - **Gmail** — to read your inbox and send alert emails
   - **Spreadsheets** — to manage VIP contacts and the alert log
   - **Script triggers** — to run automatically on a schedule

> **Note:** You may see a "This app isn't verified" warning. Click **Advanced → Go to VIP Priority Alert (unsafe)** to proceed. This is normal for personal scripts.

### Step 4: Add Your VIPs

1. Go to the **⭐ VIP Contacts** sheet (created automatically)
2. Fill in your important contacts:

| Name | Email | Priority | Alert Method |
|------|-------|----------|--------------|
| Jane Smith | jane@company.com | 🔴 Critical | 📧 Email Alert |
| Bob Lee | bob@client.com | 🟡 Important | 📧 Email Alert |
| Updates Bot | updates@tool.com | 🟢 Normal | ⭐ Star + Label Only |

- **Priority** and **Alert Method** use dropdown menus — just click the cell
- Add as many VIPs as you need

### Step 5: Configure Settings

1. Click **🕷 TAKScripts → ⚙️ Settings**
2. Set your preferences:
   - **Check Frequency** — how often to scan (default: every 5 minutes)
   - **Alert Email** — where to send notifications (defaults to your Gmail)
   - **Keywords** — subject line words that trigger alerts regardless of sender
   - **Quiet Hours** — suppress email alerts during off-hours (emails still get starred)
3. Click **Save Settings**

### Step 6: Start Monitoring

1. Click **🕷 TAKScripts → ▶️ Start Monitoring**
2. The script will confirm it is active and show your VIP count

That's it. You're live.

---

## Menu Reference

| Menu Item | What It Does |
|-----------|-------------|
| ⚙️ Settings | Opens the settings sidebar |
| ▶️ Start Monitoring | Activates the script and installs the trigger |
| ⏹ Stop Monitoring | Deactivates the script and removes the trigger |
| 🧪 Test Run | Scans last 10 emails and shows what would trigger — no alerts sent |
| 📊 View Alert Log | Jumps to the alert log sheet |
| ℹ️ About TAKScripts | Shows version and brand info |

---

## How It Works

1. A time-based trigger runs `checkForVIPEmails` on your chosen interval
2. The script searches for unread inbox emails from the last 24 hours
3. Each email is checked against your VIP list and keyword triggers
4. Matches are processed based on their alert method:
   - **Email Alert** — sends a branded HTML notification to your alert address
   - **Star + Label Only** — silently stars and labels the email in Gmail
5. All VIP emails get the **⭐ VIP** label in Gmail automatically
6. Every alert is logged to the **📋 Alert Log** sheet with timestamp and details
7. Processed message IDs are tracked to prevent duplicate alerts

---

## Priority Levels

| Level | When to Use |
|-------|-------------|
| 🔴 Critical | Emails you need to see within minutes (boss, key clients) |
| 🟡 Important | Emails that matter but can wait an hour (team leads, partners) |
| 🟢 Normal | Emails worth tracking but not urgent (regular contacts) |

---

## Keyword Triggers

Add comma-separated keywords in Settings. If any keyword appears in an email subject, it triggers an alert at 🟡 Important priority — even if the sender is not on your VIP list.

**Example keywords:** `urgent, asap, time-sensitive, action required, deadline`

---

## Quiet Hours

When enabled, quiet hours suppress email alert notifications during the specified window (e.g., 10 PM to 7 AM). VIP emails received during quiet hours are still starred and labeled — you just won't get pinged.

---

## Troubleshooting

**Menu doesn't appear?**
Reload the spreadsheet. The `onOpen` function runs when the sheet loads.

**Not getting alerts?**
- Check that monitoring is active (▶️ Start Monitoring)
- Verify your alert email in Settings
- Make sure VIP emails are in the correct format (e.g., `user@domain.com`)
- Run 🧪 Test Run to see if emails are matching

**Getting duplicate alerts?**
This shouldn't happen — the script tracks processed message IDs. If it does, try ⏹ Stop then ▶️ Start to reset the tracking.

**Trigger quota errors?**
Google limits Apps Script triggers. If you hit limits, increase the check frequency to 10 or 15 minutes.

---

## Permissions Used

| Permission | Why |
|-----------|-----|
| Gmail (read, compose, labels) | Read inbox, send alerts, apply VIP label |
| Spreadsheets | Store VIP contacts and alert log |
| Script Properties | Save settings and processed message IDs |
| Triggers | Run automatically on a schedule |

---

## Uninstall

1. Click **🕷 TAKScripts → ⏹ Stop Monitoring** to remove triggers
2. In the spreadsheet, go to **Extensions → Apps Script**
3. Click the **Triggers** icon (clock) in the left sidebar
4. Delete any remaining triggers
5. Delete the spreadsheet if no longer needed

---

*Built by [TAKScripts](https://takscripts.store) — Pre-built Google Apps Scripts for small business.*
