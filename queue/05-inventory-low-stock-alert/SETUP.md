# Inventory Low-Stock Alert — Setup Guide

**By TAKScripts (TAK Ventures)**
Version 1.0

---

## What It Does

Monitors your product inventory in Google Sheets and sends branded email alerts when items drop below their reorder level. Automatically color-codes stock statuses, tracks alert history, and can optionally email suppliers with reorder requests.

---

## Quick Start (5 minutes)

### 1. Create the Spreadsheet

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Name it something like **"Inventory Tracker"**.

### 2. Add the Script

1. Go to **Extensions > Apps Script**.
2. Delete any code in the editor.
3. Paste the entire contents of `inventory-low-stock-alert.gs`.
4. Click the **Save** icon (or Ctrl+S / Cmd+S).
5. Name the project **"Inventory Low-Stock Alert"**.

### 3. Run Initial Setup

1. In the Apps Script editor, select `initialSetup` from the function dropdown.
2. Click **Run**.
3. When prompted, click **Review Permissions** > choose your Google account > **Allow**.
4. The script will create two sheets with sample data:
   - **📦 Inventory** — your product list with dashboard
   - **📊 Alert History** — log of all alerts sent

### 4. Configure Settings

1. Go back to your spreadsheet.
2. Click **🕷 TAKScripts > ⚙️ Settings** in the menu bar.
3. Set your **alert email address**.
4. Choose your **check frequency**:
   - **On Edit** — checks instantly when you change stock values
   - **Every Hour** — runs automatically every hour
   - **Daily** — runs once per day at 8 AM
5. Optionally enable **Auto-Email Suppliers** to send reorder requests.
6. Click **Save Settings**.

---

## Sheet Layout

### 📦 Inventory Sheet

| Column | Description |
|--------|-------------|
| **Product Name** | Name of the product |
| **SKU** | Stock keeping unit / product code |
| **Category** | Product category for organization |
| **Current Stock** | Current quantity on hand (edit this!) |
| **Reorder Level** | Alert when stock falls to or below this number |
| **Supplier** | Supplier name |
| **Supplier Email** | Supplier's email for auto-reorder requests |
| **Status** | Auto-updated: In Stock / Low Stock / Out of Stock |

**Row 1** is a dashboard showing total products, in stock, low stock, and out of stock counts.

**Row 2** is the header row (frozen).

**Row 3+** is your product data.

### 📊 Alert History Sheet

Automatically logs every alert with:
- Timestamp
- Product Name
- SKU
- Stock Level
- Reorder Level
- Action Taken

---

## Menu Options

Access everything from **🕷 TAKScripts** in the menu bar:

| Menu Item | What It Does |
|-----------|-------------|
| ⚙️ Settings | Opens the settings sidebar |
| ▶️ Check Stock Now | Runs a full stock check and sends alerts |
| 🧪 Test Run | Shows what would happen without sending any emails |
| 📊 View Alert History | Jumps to the alert history sheet |
| ℹ️ About TAKScripts | Shows version and brand info |

---

## How Alerts Work

1. The script compares **Current Stock** against **Reorder Level** (multiplied by your threshold setting).
2. Items at or below the threshold are flagged as **Low Stock**.
3. Items at zero are flagged as **Out of Stock**.
4. A branded HTML email is sent to your alert email with a table of all flagged items.
5. If supplier emails are enabled, grouped reorder request emails are sent to each supplier.
6. Every alert is logged to the Alert History sheet.

### Threshold Multiplier

The threshold multiplier lets you get alerted earlier:

| Multiplier | Reorder Level = 20 | Alert When Stock Hits |
|-----------|--------------------|-----------------------|
| 1.0x | 20 | 20 or below |
| 1.25x | 20 | 25 or below |
| 1.5x | 20 | 30 or below |
| 2.0x | 20 | 40 or below |

---

## Permissions Required

When you first run the script, Google will ask for these permissions:

- **Google Sheets** — to read/write inventory data
- **Gmail** — to send alert emails
- **Script triggers** — to run scheduled checks

These permissions are standard for Apps Scripts that work with Sheets and email. The script only accesses your inventory spreadsheet and sends emails to the addresses you configure.

---

## Troubleshooting

**Menu not showing?**
Reload the spreadsheet. The menu appears on every open.

**No alerts being sent?**
- Check that your alert email is set in Settings.
- Make sure at least one product has Current Stock at or below its Reorder Level.
- Run **🕷 TAKScripts > 🧪 Test Run** to see what the script detects.

**Trigger not firing?**
- Go to Apps Script > Triggers (clock icon) to verify your trigger exists.
- Re-save your settings to recreate the trigger.

**Permission errors?**
- Re-run `initialSetup` from the Apps Script editor to re-authorize.

---

## Customization Tips

- **Add more products**: Just add rows below the existing data. The script reads all rows from row 3 onward.
- **Change categories**: The Category column is free-form — use whatever makes sense for your business.
- **Multiple alert recipients**: Put comma-separated emails in the alert email field.
- **Remove sample data**: Delete the sample rows and add your own products.

---

## Support

Part of the **TAKScripts** collection.
Pre-built Google Apps Scripts for small business.

[takscripts.store](https://takscripts.store)
