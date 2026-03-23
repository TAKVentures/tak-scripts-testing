# Follow-Up Nudger — Setup Guide

**By TAKScripts · takscripts.store**

Never let an important email slip through the cracks. Follow-Up Nudger scans your sent mail for emails that haven't received a reply, tracks them in a Google Sheet, and optionally sends polite follow-up nudges automatically.

---

## Quick Start (5 minutes)

### 1. Create a Google Sheet

- Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet
- Name it something like **"Follow-Up Nudger"**

### 2. Add the Script

- In your spreadsheet, go to **Extensions > Apps Script**
- Delete any code in the editor
- Paste the entire contents of `follow-up-nudger.gs`
- Click the **Save** icon (or Ctrl+S / Cmd+S)
- Name the project **"Follow-Up Nudger"**

### 3. Authorize Permissions

- Close the Apps Script editor and **reload your spreadsheet**
- You'll see a new menu: **🕷 TAKScripts**
- Click **🕷 TAKScripts > ⚙️ Settings**
- Google will ask you to authorize the script — click through the prompts:
  - "This app isn't verified" > **Advanced** > **Go to Follow-Up Nudger (unsafe)**
  - Review the permissions and click **Allow**

### 4. Configure Settings

The settings sidebar will open. Configure:

| Setting | Default | Description |
|---------|---------|-------------|
| Days Before Follow-Up | 3 | How many days to wait before flagging an email |
| Max Follow-Ups | 2 | Maximum follow-up emails per thread |
| Auto-Send Follow-Ups | OFF | Automatically send follow-up emails |
| Email Digest | ON | Receive a summary email of pending follow-ups |
| Subject Line | Re: {{subject}} | Subject for follow-up emails |
| Message Template | (polite nudge) | Body of follow-up emails |
| Exclude Domains | noreply, newsletter... | Skip emails to these addresses |

Click **Save Settings** when done.

### 5. Test It

- Click **🕷 TAKScripts > 🧪 Test Run**
- This scans your sent mail and shows what would be flagged — **no emails are sent**
- Review the results to make sure the script is picking up the right emails

### 6. Start It

- Click **🕷 TAKScripts > ▶️ Start Follow-Up Nudger**
- The script will run immediately and then check every hour going forward
- A "📬 Follow-Up Tracker" sheet tab will appear with your pending follow-ups

---

## Features

### Tracker Sheet

The **📬 Follow-Up Tracker** tab shows all emails awaiting a reply:

| Column | Description |
|--------|-------------|
| Date Sent | When you originally sent the email |
| To | The recipient |
| Subject | Email subject line |
| Days Waiting | How many days since you sent it |
| Status | Pending (green), Overdue (yellow), or Urgent (red) |
| Follow-Up Sent | Whether a follow-up was sent and how many |
| Thread ID | Gmail thread identifier (for internal tracking) |

Status colors:
- **Green (Pending)**: Less than 5 days, still within normal range
- **Yellow (Overdue)**: 5-6 days, getting stale
- **Red (Urgent)**: 7+ days, needs attention

### Auto-Send Follow-Ups

When enabled, the script will automatically send a polite follow-up email using your configured template. The follow-up is sent as a reply to the original thread.

Template variables:
- `{{name}}` — Recipient's first name
- `{{subject}}` — Original email subject
- `{{days}}` — Number of days since you sent the email

### Email Digest

When enabled, you'll receive a branded HTML email summarizing all pending follow-ups. This is sent to your own email address each time the script runs and finds pending items.

### Activity Log

The **📝 Follow-Up Log** tab records all actions taken by the script:
- Settings changes
- Follow-up emails sent
- Digest emails sent
- Errors encountered
- Test runs

---

## Template Variables

Use these in your follow-up message and subject line:

| Variable | Replaced With | Example |
|----------|--------------|---------|
| `{{name}}` | Recipient's first name | "Sarah" |
| `{{subject}}` | Original email subject | "Q1 Report Review" |
| `{{days}}` | Days since you sent it | "5" |

### Example Template

**Subject:** `Re: {{subject}}`

**Message:**
```
Hi {{name}},

I wanted to follow up on my email regarding "{{subject}}" sent {{days}} days ago. I understand you may be busy, but I wanted to make sure this didn't slip through the cracks.

Please let me know if you have any questions or if there's a better time to connect.

Thanks,
[Your Name]
```

---

## Menu Reference

| Menu Item | What It Does |
|-----------|-------------|
| ⚙️ Settings | Opens the settings sidebar |
| ▶️ Start Follow-Up Nudger | Creates the hourly trigger and runs immediately |
| ⏹ Stop Follow-Up Nudger | Removes the trigger (preserves sheet data) |
| 🧪 Test Run | Scans sent mail and shows results without sending anything |
| 📊 View Log | Navigates to the activity log sheet |
| ℹ️ About TAKScripts | Shows version and brand info |

---

## Permissions Required

The script needs these Google permissions:

| Permission | Why |
|-----------|-----|
| Gmail (read/send) | To scan sent mail and send follow-ups |
| Sheets (read/write) | To create and update the tracker and log sheets |
| Script Properties | To store settings and tracking state |
| Triggers | To run automatically on a schedule |

---

## FAQ

**Q: Will it send follow-ups to every email I've ever sent?**
No. It only scans the last 14 days of sent mail (or longer if your follow-up threshold is very high). It also respects your exclude list and max follow-ups setting.

**Q: What if someone replies after a follow-up was sent?**
The script detects replies and automatically removes that thread from the pending list on the next run.

**Q: Can I use it without auto-send?**
Absolutely. Turn off "Auto-Send Follow-Ups" and turn on "Email Digest" to just get notified about emails that need attention. You can then follow up manually.

**Q: Will it create duplicate follow-ups?**
No. The script is idempotent — it tracks which follow-ups have already been sent using unique thread ID + count keys stored in script properties.

**Q: How do I stop it?**
Click **🕷 TAKScripts > ⏹ Stop Follow-Up Nudger**. This removes the trigger but keeps all your sheet data intact.

**Q: Can I change settings while it's running?**
Yes. Open Settings, make changes, and save. The next run will use the new settings.

---

## Troubleshooting

| Issue | Solution |
|-------|---------|
| Menu doesn't appear | Reload the spreadsheet. The menu loads on open. |
| "Authorization required" error | Click through the auth prompts. See Step 3 above. |
| No emails showing up | Check your "Days Before Follow-Up" setting. Emails newer than this threshold won't appear. |
| Too many false positives | Add common automated senders to your exclude list. |
| Script not running | Go to Extensions > Apps Script > Triggers to verify the trigger exists. |

---

## Support

Visit [takscripts.store](https://takscripts.store) for more scripts and support.

*Powered by TAKScripts · takscripts.store*
