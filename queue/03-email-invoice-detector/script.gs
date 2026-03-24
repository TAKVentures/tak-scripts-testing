/**
 * Email Invoice Detector by TAKScripts
 * =====================================
 * Automatically scans your Gmail inbox for invoices, receipts, and payment
 * confirmations, then logs them to a branded expense dashboard.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Expense dashboard: This Month total, All Time total, Largest invoice,
 *   Top category, and total count — always up to date
 * - Detects invoices, receipts, payment confirmations, order confirmations
 * - Extracts sender, date, subject, detected amount, and category
 * - Color-coded category badges on every row
 * - Configurable scan frequency, Gmail label, amount threshold alerts
 * - Runs on a schedule (time-driven trigger) or manually
 * - Idempotent — tracks processed message IDs to prevent duplicates
 * - Test Run mode shows detections without logging
 * - HTML-formatted alert emails for high-amount invoices
 * - Refresh Stats recalculates dashboard totals at any time
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  medGray: '#666666',
  border: '#E0E0E0',
  successBg: '#E8F5E9', successText: '#2E7D32',
  warningBg: '#FFF8E1', warningText: '#F57F17',
  errorBg: '#FFEBEE',   errorText: '#C62828',
  infoBg: '#E3F2FD',    infoText: '#1565C0',
  headerFont: 'Roboto Mono',
  bodyFont: 'Roboto',
  footerText: 'Powered by TAKScripts \u00B7 takscripts.store',
};

// Dashboard layout: rows 1-2 = stats bar, row 3 = column headers, row 4+ = data
const DASHBOARD_HEADER_ROW = 3;

const DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Expense Dashboard';
const PROP_SETTINGS = 'eid_settings';
const PROP_PROCESSED = 'eid_processed_ids';

// Column index map (1-based)
const COL = {
  date:           1,
  sender:         2,
  subject:        3,
  amount:         4,
  amountNumeric:  5,
  category:       6,
  keyword:        7,
  messageId:      8,
};

const LOG_HEADERS = [
  'Date', 'Sender', 'Subject', 'Amount',
  'Amount (Numeric)', 'Category', 'Matched Keyword', 'Message ID',
];

/**
 * Keywords used to identify invoices and receipts, grouped by category.
 */
const DETECTION_RULES = {
  'Invoice':              ['invoice', 'inv #', 'inv#', 'invoice number', 'invoice attached', 'invoice enclosed'],
  'Receipt':              ['receipt', 'your receipt', 'payment receipt', 'transaction receipt', 'e-receipt'],
  'Payment Confirmation': ['payment confirmation', 'payment received', 'payment successful', 'payment processed', 'payment accepted'],
  'Order Confirmation':   ['order confirmation', 'order confirmed', 'order #', 'order number', 'order summary', 'your order'],
  'Subscription':         ['subscription', 'recurring payment', 'renewal', 'monthly charge', 'annual charge', 'billing cycle'],
  'Refund':               ['refund', 'refunded', 'credit issued', 'money back'],
};

/**
 * Amount regex patterns — supports $, USD, EUR, GBP, £, €.
 */
const AMOUNT_PATTERNS = [
  /\$\s?[\d,]+\.?\d{0,2}/g,
  /USD\s?[\d,]+\.?\d{0,2}/gi,
  /[\d,]+\.?\d{0,2}\s?USD/gi,
  /EUR\s?[\d,]+\.?\d{0,2}/gi,
  /[\d,]+\.?\d{0,2}\s?EUR/gi,
  /GBP\s?[\d,]+\.?\d{0,2}/gi,
  /£\s?[\d,]+\.?\d{0,2}/g,
  /€\s?[\d,]+\.?\d{0,2}/g,
];

// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

/**
 * Creates the TAKScripts menu when the spreadsheet opens.
 * This runs automatically — no setup needed.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('\uD83D\uDD77 TAKScripts')
    .addItem('\u2699\uFE0F Settings', 'showSettings')
    .addItem('\u25B6\uFE0F Run Scan Now', 'runScan')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Dashboard', 'viewDashboard')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run (no logging)', 'testRun')
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml_())
    .setTitle('Invoice Detector Settings')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Email Invoice Detector</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 \u00B7 by TAK Ventures</p>' +
      '<hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">' +
      '<p style="font-size: 13px; color: #333; line-height: 1.6;">' +
        'Automatically finds invoices and receipts in your Gmail<br>' +
        'and logs them to a clean expense dashboard.<br><br>' +
        'Part of the <strong>TAKScripts</strong> collection.<br>' +
        'Pre-built Google Apps Scripts for small business.' +
      '</p>' +
      '<p style="margin-top: 16px;">' +
        '<a href="https://takscripts.store" target="_blank" ' +
           'style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">' +
          'takscripts.store \u2192' +
        '</a>' +
      '</p>' +
    '</div>'
  ).setWidth(300).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to the Expense Dashboard sheet.
 */
function viewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = getOrCreateDashboard_(ss);
  }
  ss.setActiveSheet(sheet);
}

// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Returns default settings when none are saved.
 * @return {Object} Default settings object.
 */
function getDefaultSettings_() {
  return {
    scanLabel: '',
    scanAge: '7',
    frequency: '60',
    alertThreshold: '500',
    alertEmail: Session.getActiveUser().getEmail() || '',
    enableAlerts: false,
    enableAutoTrigger: false,
    categories: Object.keys(DETECTION_RULES).join(', '),
  };
}

/**
 * Save settings from the sidebar.
 * @param {Object} settings - The settings object from the sidebar form.
 * @return {Object} Success confirmation.
 */
function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_SETTINGS, JSON.stringify(settings));
    if (settings.enableAutoTrigger) {
      installTrigger_(parseInt(settings.frequency, 10) || 60);
    } else {
      removeTrigger_();
    }
    return { success: true };
  } catch (e) {
    Logger.log('Error saving settings: ' + e.message);
    throw new Error('Failed to save settings: ' + e.message);
  }
}

/**
 * Load saved settings, falling back to defaults.
 * @return {Object} The merged settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_SETTINGS);
  const defaults = getDefaultSettings_();
  if (!raw) return defaults;
  try {
    return Object.assign({}, defaults, JSON.parse(raw));
  } catch (e) {
    Logger.log('Corrupted settings, returning defaults: ' + e.message);
    return defaults;
  }
}

// ═══════════════════════════════════════════
// DASHBOARD — THE HEART OF THE PRODUCT
// ═══════════════════════════════════════════

/**
 * Creates or gets the Expense Dashboard sheet with a stats bar and branded design.
 *
 * Layout:
 *   Row 1 — Stats values  (large gold numbers on dark background)
 *   Row 2 — Stats labels  (small uppercase on dark background)
 *   Row 3 — Column headers
 *   Row 4+ — Invoice/expense data
 *
 * Stats: THIS MONTH | ALL TIME | LARGEST | TOP CATEGORY | TOTAL
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The dashboard sheet.
 */
function getOrCreateDashboard_(ss) {
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);

  // ── Stats Row 1: Values ──
  sheet.getRange(1, 1, 1, 5).setValues([['$0', '$0', '$0', '—', '0']]);

  // ── Stats Row 2: Labels ──
  sheet.getRange(2, 1, 1, 5).setValues([[
    'THIS MONTH', 'ALL TIME', 'LARGEST', 'TOP CATEGORY', 'TOTAL',
  ]]);

  // Style stats values (row 1) — large bold gold
  sheet.getRange(1, 1, 1, 5)
    .setFontFamily(BRAND.headerFont)
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);

  // Style stats labels (row 2) — tiny uppercase
  sheet.getRange(2, 1, 1, 5)
    .setFontFamily(BRAND.headerFont)
    .setFontSize(8)
    .setFontWeight('bold')
    .setFontColor('#888888')
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top');
  sheet.setRowHeight(2, 22);

  // Extend dark background to all columns in stat rows
  const totalCols = LOG_HEADERS.length;
  if (totalCols > 5) {
    sheet.getRange(1, 6, 1, totalCols - 5).setBackground(BRAND.darkBg);
    sheet.getRange(2, 6, 1, totalCols - 5).setBackground(BRAND.darkBg);
  }

  // ── Column Headers (Row 3) ──
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, LOG_HEADERS.length)
    .setFontFamily(BRAND.headerFont)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(DASHBOARD_HEADER_ROW, 32);

  // Freeze through header row
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // ── Column Widths ──
  sheet.setColumnWidth(COL.date,          160);
  sheet.setColumnWidth(COL.sender,        230);
  sheet.setColumnWidth(COL.subject,       320);
  sheet.setColumnWidth(COL.amount,        110);
  sheet.setColumnWidth(COL.amountNumeric, 130);
  sheet.setColumnWidth(COL.category,      160);
  sheet.setColumnWidth(COL.keyword,       150);
  sheet.setColumnWidth(COL.messageId,     160);

  // Hide internal columns (Message ID used for dedup, Amount Numeric for calculations)
  sheet.hideColumns(COL.messageId);

  // Move to first tab
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  return sheet;
}

/**
 * Refreshes the stats bar (rows 1-2) by scanning all logged invoice data.
 * Computes: This Month total, All Time total, Largest single invoice,
 * Top category by frequency, and total count.
 */
function refreshDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No dashboard yet. Run a scan first.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;

  if (lastRow < dataStartRow) {
    sheet.getRange(1, 1, 1, 5).setValues([['$0', '$0', '$0', '—', '0']]);
    return;
  }

  const numDataRows = lastRow - dataStartRow + 1;
  const data = sheet.getRange(dataStartRow, 1, numDataRows, LOG_HEADERS.length).getValues();

  let thisMonth = 0;
  let allTime = 0;
  let largest = 0;
  let largestFormatted = '$0';
  let total = 0;
  const categoryCounts = {};

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.date - 1]) continue;

    const dateVal = row[COL.date - 1];
    const amountNumeric = parseFloat(row[COL.amountNumeric - 1]) || 0;
    const amountFormatted = String(row[COL.amount - 1] || '');
    const category = String(row[COL.category - 1] || '').trim();

    total++;

    // This month
    const rowDate = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (!isNaN(rowDate.getTime()) &&
        rowDate.getMonth() === currentMonth &&
        rowDate.getFullYear() === currentYear) {
      thisMonth += amountNumeric;
    }

    allTime += amountNumeric;

    if (amountNumeric > largest) {
      largest = amountNumeric;
      largestFormatted = amountFormatted || ('$' + amountNumeric.toFixed(2));
    }

    if (category) {
      categoryCounts[category] = (categoryCounts[category] || 0) + 1;
    }
  }

  // Find top category
  let topCategory = '—';
  let topCount = 0;
  for (const cat in categoryCounts) {
    if (categoryCounts[cat] > topCount) {
      topCount = categoryCounts[cat];
      topCategory = cat;
    }
  }

  // Format currency values
  const fmt = function(n) {
    return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  };

  sheet.getRange(1, 1, 1, 5).setValues([[
    fmt(thisMonth),
    fmt(allTime),
    largestFormatted || '$0',
    topCategory,
    String(total),
  ]]);
}

// ═══════════════════════════════════════════
// TRIGGER MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Installs a time-driven trigger for automatic scanning.
 * Removes any existing trigger first to avoid duplicates.
 * @param {number} intervalMinutes - How often to run (in minutes).
 */
function installTrigger_(intervalMinutes) {
  removeTrigger_();
  const validIntervals = [1, 5, 10, 15, 30];
  if (validIntervals.includes(intervalMinutes)) {
    ScriptApp.newTrigger('runScan').timeBased().everyMinutes(intervalMinutes).create();
  } else {
    const hours = Math.max(1, Math.round(intervalMinutes / 60));
    ScriptApp.newTrigger('runScan').timeBased().everyHours(hours).create();
  }
  Logger.log('Trigger installed: every ' + intervalMinutes + ' minutes.');
}

/**
 * Removes any existing runScan triggers.
 */
function removeTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runScan') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main scan function — searches Gmail for invoices/receipts and logs them.
 * Safe to call from a trigger or manually from the menu.
 */
function runScan() {
  const settings = loadSettings();
  const results = scanEmails_(settings, false);

  try {
    const ui = SpreadsheetApp.getUi();
    if (results.total === 0) {
      ui.alert('\u2705 Scan Complete\n\nNo new invoices or receipts found.');
    } else {
      ui.alert(
        '\u2705 Scan Complete\n\n' +
        'Found ' + results.total + ' new invoice/receipt email(s).\n' +
        'Logged to the Expense Dashboard.\n\n' +
        (results.alerts > 0
          ? '\u26A0\uFE0F ' + results.alerts + ' item(s) exceeded your amount threshold.'
          : 'No threshold alerts triggered.')
      );
    }
  } catch (e) {
    Logger.log('Scan complete. Found ' + results.total + ' new items.');
  }
}

/**
 * Test Run — shows what the scanner would detect without logging anything.
 */
function testRun() {
  const settings = loadSettings();
  const results = scanEmails_(settings, true);

  const lines = ['\uD83E\uDDEA TEST RUN \u2014 Nothing was logged', ''];

  if (results.items.length === 0) {
    lines.push('No invoices or receipts detected in the scanned emails.');
    lines.push('');
    lines.push('Searched: ' + results.threadsScanned + ' thread(s) from the last ' + (settings.scanAge || 7) + ' day(s).');
  } else {
    lines.push('Detected ' + results.items.length + ' invoice/receipt email(s):');
    lines.push('\u2500'.repeat(40));
    for (const item of results.items) {
      lines.push('');
      lines.push('From: ' + item.sender);
      lines.push('Subject: ' + item.subject);
      lines.push('Category: ' + item.category);
      lines.push('Amount: ' + (item.amount || 'Not detected'));
      lines.push('Date: ' + item.date);
    }
  }

  try {
    SpreadsheetApp.getUi().alert(lines.join('\n'));
  } catch (e) {
    Logger.log('Scan finished. Detected: ' + results.items.length);
  }
}

/**
 * Core scanning logic. Processes Gmail threads and detects invoices.
 * @param {Object} settings - Current settings.
 * @param {boolean} dryRun - If true, detect but don't log or send alerts.
 * @return {Object} Results summary { total, alerts, items, threadsScanned }.
 */
function scanEmails_(settings, dryRun) {
  const age = parseInt(settings.scanAge, 10) || 7;
  let query = 'newer_than:' + age + 'd';
  if (settings.scanLabel && settings.scanLabel.trim()) {
    query += ' label:' + settings.scanLabel.trim();
  } else {
    query += ' in:inbox';
  }

  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_PROCESSED) || '[]';
  let processedIds;
  try {
    processedIds = new Set(JSON.parse(processedRaw));
  } catch (e) {
    processedIds = new Set();
  }

  const enabledCategories = (settings.categories || Object.keys(DETECTION_RULES).join(', '))
    .split(',').map(function(c) { return c.trim(); }).filter(Boolean);

  const threshold = parseFloat(settings.alertThreshold) || 0;
  const results = { total: 0, alerts: 0, items: [], threadsScanned: 0 };

  let threads;
  try {
    threads = GmailApp.search(query, 0, 200);
  } catch (e) {
    Logger.log('Gmail search error: ' + e.message);
    return results;
  }

  results.threadsScanned = threads.length;

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const messageId = message.getId();
      if (processedIds.has(messageId)) continue;

      const subject = message.getSubject() || '';
      const from = message.getFrom() || '';
      const date = message.getDate();
      let body = '';
      try {
        body = message.getPlainBody() || '';
      } catch (e) {
        body = '';
      }

      const detection = detectInvoice_(subject, body, enabledCategories);
      if (!detection.matched) continue;

      const amount = extractAmount_(subject, body);
      const item = {
        messageId:     messageId,
        sender:        from,
        senderEmail:   extractEmail_(from),
        subject:       subject,
        date:          formatDate_(date),
        rawDate:       date,
        category:      detection.category,
        amount:        amount.formatted,
        amountNumeric: amount.numeric,
        matchedKeyword: detection.keyword,
      };

      results.items.push(item);
      results.total++;

      if (!dryRun) {
        logInvoice_(item);
        processedIds.add(messageId);

        if (settings.enableAlerts && threshold > 0 && amount.numeric >= threshold) {
          results.alerts++;
          sendThresholdAlert_(item, threshold, settings.alertEmail);
        }
      }
    }
  }

  if (!dryRun && results.total > 0) {
    const trimmed = [...processedIds].slice(-5000);
    props.setProperty(PROP_PROCESSED, JSON.stringify(trimmed));
    refreshDashboardStats();
  }

  Logger.log('Scan finished. Detected: ' + results.total + ', Alerts: ' + results.alerts);
  return results;
}

/**
 * Checks if a message subject or body matches invoice/receipt keywords.
 * @param {string} subject - Email subject.
 * @param {string} body - Email plain text body.
 * @param {string[]} enabledCategories - Which categories to check.
 * @return {Object} { matched: boolean, category: string, keyword: string }
 */
function detectInvoice_(subject, body, enabledCategories) {
  const searchText = (subject + ' ' + body).toLowerCase();
  for (const category of enabledCategories) {
    const keywords = DETECTION_RULES[category];
    if (!keywords) continue;
    for (const keyword of keywords) {
      if (searchText.includes(keyword.toLowerCase())) {
        return { matched: true, category: category, keyword: keyword };
      }
    }
  }
  return { matched: false, category: '', keyword: '' };
}

/**
 * Extracts the highest dollar/currency amount from subject and body text.
 * @param {string} subject - Email subject.
 * @param {string} body - Email plain text body (first 3000 chars used).
 * @return {Object} { formatted: string, numeric: number }
 */
function extractAmount_(subject, body) {
  const searchText = subject + ' ' + body.substring(0, 3000);
  let highestAmount = 0;
  let formattedAmount = '';

  for (const pattern of AMOUNT_PATTERNS) {
    pattern.lastIndex = 0;
    const matches = searchText.match(pattern);
    if (!matches) continue;
    for (const match of matches) {
      const cleaned = match.replace(/[^0-9.,]/g, '').replace(/,/g, '');
      const value = parseFloat(cleaned);
      if (!isNaN(value) && value > highestAmount && value < 1000000) {
        highestAmount = value;
        formattedAmount = match.trim();
      }
    }
  }

  return { formatted: formattedAmount || '', numeric: highestAmount };
}

// ═══════════════════════════════════════════
// SHEET LOGGING
// ═══════════════════════════════════════════

/**
 * Logs a detected invoice to the Expense Dashboard.
 * Creates the dashboard if it doesn't exist.
 * Data rows start at DASHBOARD_HEADER_ROW + 1.
 * @param {Object} item - The detected invoice data.
 */
function logInvoice_(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDashboard_(ss);

  const lastRow = sheet.getLastRow();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const newRow = Math.max(lastRow + 1, dataStartRow);

  const rowData = [
    item.rawDate,
    item.senderEmail,
    item.subject,
    item.amount || '\u2014',
    item.amountNumeric || 0,
    item.category,
    item.matchedKeyword,
    item.messageId,
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

  // Alternating row colors
  const bg = newRow % 2 === 0 ? BRAND.lightGray : BRAND.white;
  sheet.getRange(newRow, 1, 1, rowData.length)
    .setBackground(bg)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setFontColor('#333333')
    .setVerticalAlignment('middle');

  // Date format
  sheet.getRange(newRow, COL.date).setNumberFormat('yyyy-mm-dd');

  // Numeric amount format
  sheet.getRange(newRow, COL.amountNumeric).setNumberFormat('$#,##0.00');

  // Category badge
  applyCategoryStyle_(sheet, newRow, COL.category, item.category);
}

/**
 * Applies a color-coded background to a category cell.
 * @param {Sheet} sheet - The log sheet.
 * @param {number} row - Row number.
 * @param {number} col - Column number.
 * @param {string} category - The invoice category.
 */
function applyCategoryStyle_(sheet, row, col, category) {
  const styles = {
    'Invoice':              { bg: BRAND.warningBg, fg: BRAND.warningText },
    'Receipt':              { bg: BRAND.successBg, fg: BRAND.successText },
    'Payment Confirmation': { bg: BRAND.successBg, fg: BRAND.successText },
    'Order Confirmation':   { bg: BRAND.infoBg,    fg: BRAND.infoText    },
    'Subscription':         { bg: BRAND.warningBg, fg: BRAND.warningText },
    'Refund':               { bg: BRAND.errorBg,   fg: BRAND.errorText   },
  };
  const style = styles[category] || { bg: '#F5F5F5', fg: BRAND.medGray };
  sheet.getRange(row, col)
    .setBackground(style.bg)
    .setFontColor(style.fg)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════
// ALERT EMAILS
// ═══════════════════════════════════════════

/**
 * Sends a branded HTML alert email when an invoice exceeds the threshold.
 * @param {Object} item - The detected invoice data.
 * @param {number} threshold - The configured alert threshold.
 * @param {string} alertEmail - Email address to send alert to.
 */
function sendThresholdAlert_(item, threshold, alertEmail) {
  if (!alertEmail) return;

  const subject = '\u26A0\uFE0F High Invoice Alert \u2014 ' + item.amount + ' from ' + item.senderEmail;

  const htmlBody =
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; max-width: 600px; margin: 0 auto;">' +
    '<div style="background: #1A1A1A; padding: 24px; text-align: center; border-radius: 8px 8px 0 0;">' +
    '<div style="font-size: 28px; margin-bottom: 4px;">\uD83D\uDD77</div>' +
    '<div style="color: #C9A84C; font-size: 16px; font-weight: 700; letter-spacing: 1px;">TAKScripts</div>' +
    '<div style="color: #888; font-size: 11px; margin-top: 4px;">Email Invoice Detector \u00B7 Alert</div>' +
    '</div>' +
    '<div style="background: #fff; border: 1px solid #eee; border-top: none; padding: 24px;">' +
    '<h2 style="margin: 0 0 16px; font-size: 18px; color: #C62828;">\u26A0\uFE0F Invoice Exceeds Threshold</h2>' +
    '<p style="font-size: 14px; color: #333; line-height: 1.6; margin: 0 0 20px;">' +
    'An email was detected with an amount of <strong>' + escapeHtml_(item.amount) + '</strong>, ' +
    'which exceeds your alert threshold of <strong>$' + threshold.toFixed(2) + '</strong>.' +
    '</p>' +
    '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">' +
    '<tr style="border-bottom: 1px solid #f0f0f0;"><td style="padding: 10px 0; color: #888; width: 120px;">From</td><td style="padding: 10px 0; color: #333;">' + escapeHtml_(item.senderEmail) + '</td></tr>' +
    '<tr style="border-bottom: 1px solid #f0f0f0;"><td style="padding: 10px 0; color: #888;">Subject</td><td style="padding: 10px 0; color: #333;">' + escapeHtml_(item.subject) + '</td></tr>' +
    '<tr style="border-bottom: 1px solid #f0f0f0;"><td style="padding: 10px 0; color: #888;">Amount</td><td style="padding: 10px 0; color: #C62828; font-weight: 700; font-size: 16px;">' + escapeHtml_(item.amount) + '</td></tr>' +
    '<tr style="border-bottom: 1px solid #f0f0f0;"><td style="padding: 10px 0; color: #888;">Category</td><td style="padding: 10px 0; color: #333;">' + escapeHtml_(item.category) + '</td></tr>' +
    '<tr><td style="padding: 10px 0; color: #888;">Date</td><td style="padding: 10px 0; color: #333;">' + escapeHtml_(item.date) + '</td></tr>' +
    '</table></div>' +
    '<div style="background: #FAFAFA; border: 1px solid #eee; border-top: none; padding: 16px; text-align: center; border-radius: 0 0 8px 8px;">' +
    '<p style="font-size: 11px; color: #999; margin: 0;">Sent by <strong>Email Invoice Detector</strong> via ' +
    '<a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none;">TAKScripts</a></p>' +
    '</div></div>';

  try {
    GmailApp.sendEmail(alertEmail, subject,
      'High Invoice Alert: ' + item.amount + ' from ' + item.senderEmail,
      { htmlBody: htmlBody }
    );
  } catch (e) {
    Logger.log('Failed to send alert email: ' + e.message);
  }
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Extracts the email address from a "From" header field.
 * Handles: "John Doe <john@example.com>" or "john@example.com"
 * @param {string} fromField - The raw From header value.
 * @return {string} The extracted email address.
 */
function extractEmail_(fromField) {
  const match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

/**
 * Formats a Date object to a human-readable string.
 * @param {Date} date - The date to format.
 * @return {string} Formatted date string (e.g., "March 15, 2026").
 */
function formatDate_(date) {
  const months = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

/**
 * Escapes HTML special characters to prevent injection in alert emails.
 * @param {string} str - The raw string.
 * @return {string} The escaped string.
 */
function escapeHtml_(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Resets processed message IDs. Use to re-scan old emails.
 * Available from Script Editor only (not in the menu for safety).
 */
function resetProcessedIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
  Logger.log('Processed IDs cleared. Next scan will re-process all emails.');
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

/**
 * Returns the full HTML for the branded settings sidebar.
 * @return {string} Complete HTML string.
 */
function getSettingsHtml_() {
  const categoryOptions = Object.keys(DETECTION_RULES)
    .map(function(cat) {
      return '<label class="checkbox-row"><input type="checkbox" value="' + cat + '" checked /><span>' + cat + '</span></label>';
    })
    .join('');

  return `<!DOCTYPE html>
<html>
<head>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: #FAFAFA;
      color: #1A1A1A;
      font-size: 13px;
    }
    .header {
      background: #1A1A1A;
      color: white;
      padding: 20px 16px;
      text-align: center;
    }
    .header .logo { font-size: 24px; margin-bottom: 4px; }
    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }
    .header .brand { color: #C9A84C; }
    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }
    .form { padding: 16px; }
    .section-title {
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 1.5px;
      color: #C9A84C;
      margin: 20px 0 12px;
      padding-bottom: 6px;
      border-bottom: 1px solid #eee;
    }
    .section-title:first-child { margin-top: 0; }
    .field { margin-bottom: 16px; }
    .field label {
      display: block;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #666;
      margin-bottom: 6px;
    }
    .field input[type="text"],
    .field input[type="number"],
    .field input[type="email"],
    .field select {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      font-family: inherit;
      transition: border-color 0.2s, box-shadow 0.2s;
      background: #fff;
    }
    .field input:focus, .field select:focus {
      outline: none;
      border-color: #C9A84C;
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }
    .inline-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .toggle-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 0;
    }
    .toggle-label { font-size: 13px; color: #333; }
    .toggle-label .sublabel { display: block; font-size: 11px; color: #999; margin-top: 2px; }
    .toggle { position: relative; width: 40px; height: 22px; flex-shrink: 0; margin-left: 12px; }
    .toggle input { opacity: 0; width: 0; height: 0; }
    .toggle .slider {
      position: absolute; inset: 0;
      background: #ccc; border-radius: 22px; cursor: pointer; transition: background 0.2s;
    }
    .toggle .slider::before {
      content: ''; position: absolute;
      height: 16px; width: 16px; left: 3px; bottom: 3px;
      background: white; border-radius: 50%; transition: transform 0.2s;
    }
    .toggle input:checked + .slider { background: #C9A84C; }
    .toggle input:checked + .slider::before { transform: translateX(18px); }
    .checkbox-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
    .checkbox-row {
      display: flex; align-items: center; gap: 6px;
      font-size: 12px; color: #444; cursor: pointer; padding: 4px 0;
    }
    .checkbox-row input[type="checkbox"] { accent-color: #C9A84C; width: 14px; height: 14px; }
    .btn {
      width: 100%; padding: 12px; border: none; border-radius: 8px;
      font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px;
    }
    .btn-primary { background: #1A1A1A; color: #C9A84C; border: 1px solid #C9A84C; }
    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }
    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .status { text-align: center; padding: 8px; font-size: 12px; margin-top: 8px; border-radius: 6px; display: none; }
    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }
    .status.error { display: block; background: #FFEBEE; color: #C62828; }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo">\uD83D\uDD77</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">Email Invoice Detector \u00B7 Settings</div>
  </div>

  <div class="form">

    <div class="section-title">Scan Settings</div>

    <div class="field">
      <label>Gmail Label (optional)</label>
      <input type="text" id="scanLabel" placeholder="e.g. invoices (leave blank for inbox)" />
      <div class="help">Only scan emails with this label. Leave blank to scan your full inbox.</div>
    </div>

    <div class="inline-fields">
      <div class="field">
        <label>Scan Last N Days</label>
        <input type="number" id="scanAge" min="1" max="365" placeholder="7" />
      </div>
      <div class="field">
        <label>Run Every (min)</label>
        <select id="frequency">
          <option value="15">15 minutes</option>
          <option value="30">30 minutes</option>
          <option value="60" selected>1 hour</option>
          <option value="360">6 hours</option>
          <option value="1440">Daily</option>
        </select>
      </div>
    </div>

    <div class="toggle-row">
      <div class="toggle-label">
        Auto-scan on schedule
        <span class="sublabel">Runs automatically in the background</span>
      </div>
      <div class="toggle">
        <input type="checkbox" id="enableAutoTrigger" />
        <span class="slider"></span>
      </div>
    </div>

    <div class="section-title">Amount Alerts</div>

    <div class="toggle-row">
      <div class="toggle-label">
        Alert me on large invoices
        <span class="sublabel">Sends you an email above the threshold</span>
      </div>
      <div class="toggle">
        <input type="checkbox" id="enableAlerts" />
        <span class="slider"></span>
      </div>
    </div>

    <div class="inline-fields">
      <div class="field">
        <label>Alert Threshold ($)</label>
        <input type="number" id="alertThreshold" min="0" placeholder="500" />
      </div>
      <div class="field">
        <label>Alert Email</label>
        <input type="email" id="alertEmail" placeholder="you@email.com" />
      </div>
    </div>

    <div class="section-title">Categories to Detect</div>

    <div class="field">
      <div class="checkbox-grid">
        ${categoryOptions}
      </div>
      <div class="help" style="margin-top: 8px;">Uncheck categories you don't want to track.</div>
    </div>

    <div id="status" class="status"></div>
    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
  </div>

  <script>
    google.script.run.withSuccessHandler(function(s) {
      document.getElementById('scanLabel').value = s.scanLabel || '';
      document.getElementById('scanAge').value = s.scanAge || 7;
      document.getElementById('frequency').value = s.frequency || '60';
      document.getElementById('enableAutoTrigger').checked = s.enableAutoTrigger || false;
      document.getElementById('enableAlerts').checked = s.enableAlerts || false;
      document.getElementById('alertThreshold').value = s.alertThreshold || 500;
      document.getElementById('alertEmail').value = s.alertEmail || '';

      // Restore category checkboxes
      var enabled = (s.categories || '').split(',').map(function(c) { return c.trim(); });
      document.querySelectorAll('.checkbox-row input[type="checkbox"]').forEach(function(cb) {
        cb.checked = enabled.includes(cb.value);
      });
    }).loadSettings();

    function save() {
      var checkedCats = [];
      document.querySelectorAll('.checkbox-row input[type="checkbox"]:checked').forEach(function(cb) {
        checkedCats.push(cb.value);
      });

      var settings = {
        scanLabel: document.getElementById('scanLabel').value.trim(),
        scanAge: document.getElementById('scanAge').value || '7',
        frequency: document.getElementById('frequency').value,
        enableAutoTrigger: document.getElementById('enableAutoTrigger').checked,
        enableAlerts: document.getElementById('enableAlerts').checked,
        alertThreshold: document.getElementById('alertThreshold').value || '500',
        alertEmail: document.getElementById('alertEmail').value.trim(),
        categories: checkedCats.join(', '),
      };

      var statusEl = document.getElementById('status');
      var saveBtn = document.getElementById('saveBtn');
      saveBtn.disabled = true;
      saveBtn.textContent = 'Saving\u2026';
      google.script.run
        .withSuccessHandler(function() {
          statusEl.textContent = '\u2713 Settings saved successfully';
          statusEl.className = 'status success';
          saveBtn.textContent = '\u2713 Saved!';
          setTimeout(function() {
            saveBtn.textContent = 'Save Settings';
            saveBtn.disabled = false;
          }, 2500);
        })
        .withFailureHandler(function(err) {
          statusEl.textContent = '\u2715 Error: ' + err.message;
          statusEl.className = 'status error';
          saveBtn.textContent = 'Save Settings';
          saveBtn.disabled = false;
        })
        .saveSettings(settings);
    }
  </script>
</body>
</html>`;
}
