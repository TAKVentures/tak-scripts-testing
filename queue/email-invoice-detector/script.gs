/**
 * Email Invoice Detector by TAKScripts
 * =====================================
 * Automatically scans your Gmail inbox for invoices, receipts, and payment
 * confirmations, then logs them to a branded Google Sheet.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Detects invoices, receipts, payment confirmations, order confirmations
 * - Extracts sender, date, subject, detected amount, and category
 * - Logs everything to a branded "📊 Invoice Log" sheet
 * - Configurable scan frequency, Gmail label, amount threshold alerts, categories
 * - Runs on a schedule (time-driven trigger) or manually
 * - Idempotent — tracks processed message IDs to prevent duplicates
 * - Test Run mode shows detections without logging
 * - HTML-formatted alert emails for high-amount invoices
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const SCRIPT_NAME = 'Email Invoice Detector';
const LOG_SHEET_NAME = '📊 Invoice Log';
const SETTINGS_SHEET_NAME = '⚙️ Settings';
const PROP_SETTINGS = 'eid_settings';
const PROP_PROCESSED = 'eid_processed_ids';

/**
 * Keywords used to identify invoices and receipts.
 * Grouped by category for classification.
 */
const DETECTION_RULES = {
  'Invoice': ['invoice', 'inv #', 'inv#', 'invoice number', 'invoice attached', 'invoice enclosed'],
  'Receipt': ['receipt', 'your receipt', 'payment receipt', 'transaction receipt', 'e-receipt'],
  'Payment Confirmation': ['payment confirmation', 'payment received', 'payment successful', 'payment processed', 'payment accepted'],
  'Order Confirmation': ['order confirmation', 'order confirmed', 'order #', 'order number', 'order summary', 'your order'],
  'Subscription': ['subscription', 'recurring payment', 'renewal', 'monthly charge', 'annual charge', 'billing cycle'],
  'Refund': ['refund', 'refunded', 'credit issued', 'money back'],
};

/**
 * Amount regex patterns to detect prices in email subjects and bodies.
 * Supports: $1,234.56 | USD 1234.56 | 1,234.56 USD | EUR 50.00
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
  ui.createMenu('🕷 TAKScripts')
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('▶️ Run Scan Now', 'runScan')
    .addSeparator()
    .addItem('🧪 Test Run (no logging)', 'testRun')
    .addItem('📊 View Invoice Log', 'viewLog')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
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
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Segoe UI', system-ui, sans-serif; padding: 20px; text-align: center;">
      <div style="font-size: 32px; margin-bottom: 8px;">🕷</div>
      <h2 style="margin: 0 0 4px; font-size: 18px;">Email Invoice Detector</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 · by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
        Automatically finds invoices and receipts in your Gmail
        and logs them to a clean, organized spreadsheet.<br><br>
        Part of the <strong>TAKScripts</strong> collection.<br>
        Pre-built Google Apps Scripts for small business.
      </p>
      <p style="margin-top: 16px;">
        <a href="https://takscripts.store" target="_blank"
           style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">
          takscripts.store →
        </a>
      </p>
    </div>
  `).setWidth(300).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to the invoice log sheet, or alerts if none exists yet.
 */
function viewLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No invoice log yet.\n\nRun a scan from the TAKScripts menu and detected invoices will appear here.'
    );
    return;
  }
  ss.setActiveSheet(sheet);
}


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Returns default settings when none are saved.
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
 * @returns {Object} Success confirmation.
 */
function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_SETTINGS, JSON.stringify(settings));

    // Manage auto-trigger based on settings
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
 * @returns {Object} The merged settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_SETTINGS);
  const defaults = getDefaultSettings_();
  if (!raw) return defaults;

  try {
    const saved = JSON.parse(raw);
    // Merge with defaults so new fields always exist
    return Object.assign({}, defaults, saved);
  } catch (e) {
    Logger.log('Corrupted settings, returning defaults: ' + e.message);
    return defaults;
  }
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

  // Apps Script only supports specific intervals: 1, 5, 10, 15, 30, 60
  const validIntervals = [1, 5, 10, 15, 30];
  if (validIntervals.includes(intervalMinutes)) {
    ScriptApp.newTrigger('runScan')
      .timeBased()
      .everyMinutes(intervalMinutes)
      .create();
  } else {
    // For 60+ minutes, use everyHours
    const hours = Math.max(1, Math.round(intervalMinutes / 60));
    ScriptApp.newTrigger('runScan')
      .timeBased()
      .everyHours(hours)
      .create();
  }

  Logger.log('Trigger installed: every ' + intervalMinutes + ' minutes.');
}

/**
 * Removes any existing scan triggers.
 */
function removeTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runScan') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
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

  // Show summary if run manually (not from trigger)
  try {
    const ui = SpreadsheetApp.getUi();
    if (results.total === 0) {
      ui.alert('✅ Scan Complete\n\nNo new invoices or receipts found.');
    } else {
      ui.alert(
        '✅ Scan Complete\n\n' +
        'Found ' + results.total + ' new invoice/receipt email(s).\n' +
        'Logged to "' + LOG_SHEET_NAME + '" sheet.\n\n' +
        (results.alerts > 0
          ? '⚠️ ' + results.alerts + ' item(s) exceeded your amount threshold.'
          : 'No threshold alerts triggered.')
      );
    }
  } catch (e) {
    // Running from trigger — no UI available, that's fine
    Logger.log('Scan complete. Found ' + results.total + ' new items.');
  }
}

/**
 * Test Run — shows what the scanner would detect without logging anything.
 */
function testRun() {
  const settings = loadSettings();
  const results = scanEmails_(settings, true);

  const lines = [];
  lines.push('🧪 TEST RUN — Nothing was logged');
  lines.push('');

  if (results.items.length === 0) {
    lines.push('No invoices or receipts detected in the scanned emails.');
    lines.push('');
    lines.push('Searched: ' + results.threadsScanned + ' thread(s) from the last ' + (settings.scanAge || 7) + ' day(s).');
  } else {
    lines.push('Detected ' + results.items.length + ' invoice/receipt email(s):');
    lines.push('─'.repeat(40));

    for (const item of results.items) {
      lines.push('');
      lines.push('From: ' + item.sender);
      lines.push('Subject: ' + item.subject);
      lines.push('Category: ' + item.category);
      lines.push('Amount: ' + (item.amount || 'Not detected'));
      lines.push('Date: ' + item.date);
    }
  }

  const output = lines.join('\n');
  Logger.log(output);
  SpreadsheetApp.getUi().alert(output);
}

/**
 * Core scanning logic. Processes Gmail threads and detects invoices.
 *
 * @param {Object} settings - Current settings.
 * @param {boolean} dryRun - If true, detect but don't log or send alerts.
 * @returns {Object} Results summary { total, alerts, items, threadsScanned }.
 */
function scanEmails_(settings, dryRun) {
  // Build Gmail search query
  const age = parseInt(settings.scanAge, 10) || 7;
  let query = 'newer_than:' + age + 'd';
  if (settings.scanLabel && settings.scanLabel.trim()) {
    query += ' label:' + settings.scanLabel.trim();
  } else {
    query += ' in:inbox';
  }

  // Load processed message IDs to prevent duplicates
  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_PROCESSED) || '[]';
  let processedIds;
  try {
    processedIds = new Set(JSON.parse(processedRaw));
  } catch (e) {
    processedIds = new Set();
  }

  // Determine which categories to scan for
  const enabledCategories = (settings.categories || Object.keys(DETECTION_RULES).join(', '))
    .split(',')
    .map(c => c.trim())
    .filter(Boolean);

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
    const messages = thread.getMessages();

    for (const message of messages) {
      const messageId = message.getId();

      // Skip already-processed messages (idempotency)
      if (processedIds.has(messageId)) continue;

      const subject = message.getSubject() || '';
      const from = message.getFrom() || '';
      const date = message.getDate();
      let body = '';

      try {
        body = message.getPlainBody() || '';
      } catch (e) {
        // Some messages may not have a plain body
        body = '';
      }

      // Check if this message matches any detection rules
      const detection = detectInvoice_(subject, body, enabledCategories);
      if (!detection.matched) continue;

      // Extract the highest amount mentioned
      const amount = extractAmount_(subject, body);

      const senderEmail = extractEmail_(from);
      const item = {
        messageId: messageId,
        sender: from,
        senderEmail: senderEmail,
        subject: subject,
        date: formatDate_(date),
        rawDate: date,
        category: detection.category,
        amount: amount.formatted,
        amountNumeric: amount.numeric,
        matchedKeyword: detection.keyword,
      };

      results.items.push(item);
      results.total++;

      if (!dryRun) {
        // Log to sheet
        logInvoice_(item);

        // Mark as processed
        processedIds.add(messageId);

        // Check threshold for alerts
        if (settings.enableAlerts && threshold > 0 && amount.numeric >= threshold) {
          results.alerts++;
          sendThresholdAlert_(item, threshold, settings.alertEmail);
        }
      }
    }
  }

  // Persist processed IDs (keep only last 5000 to avoid property size limits)
  if (!dryRun && results.total > 0) {
    const processedArray = [...processedIds];
    const trimmed = processedArray.slice(-5000);
    props.setProperty(PROP_PROCESSED, JSON.stringify(trimmed));
  }

  Logger.log('Scan finished. Detected: ' + results.total + ', Alerts: ' + results.alerts);
  return results;
}

/**
 * Checks if a message subject or body matches invoice/receipt keywords.
 *
 * @param {string} subject - Email subject.
 * @param {string} body - Email plain text body.
 * @param {string[]} enabledCategories - Which categories to check.
 * @returns {Object} { matched: boolean, category: string, keyword: string }
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
 *
 * @param {string} subject - Email subject.
 * @param {string} body - Email plain text body (first 3000 chars used).
 * @returns {Object} { formatted: string, numeric: number }
 */
function extractAmount_(subject, body) {
  // Only scan the first 3000 chars of body for performance
  const searchText = subject + ' ' + body.substring(0, 3000);
  let highestAmount = 0;
  let formattedAmount = '';

  for (const pattern of AMOUNT_PATTERNS) {
    // Reset lastIndex for global regex
    pattern.lastIndex = 0;
    const matches = searchText.match(pattern);
    if (!matches) continue;

    for (const match of matches) {
      // Strip currency symbols and parse
      const cleaned = match.replace(/[^0-9.,]/g, '').replace(/,/g, '');
      const value = parseFloat(cleaned);
      if (!isNaN(value) && value > highestAmount && value < 1000000) {
        highestAmount = value;
        formattedAmount = match.trim();
      }
    }
  }

  return {
    formatted: formattedAmount || '',
    numeric: highestAmount,
  };
}


// ═══════════════════════════════════════════
// SHEET LOGGING
// ═══════════════════════════════════════════

/**
 * Logs a detected invoice to the branded log sheet.
 * Creates and styles the sheet if it doesn't exist.
 *
 * @param {Object} item - The detected invoice data.
 */
function logInvoice_(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = createLogSheet_(ss);
  }

  // Append the data row
  const newRow = sheet.getLastRow() + 1;
  const rowData = [
    item.rawDate,
    item.senderEmail,
    item.subject,
    item.amount || '—',
    item.amountNumeric || 0,
    item.category,
    item.matchedKeyword,
    item.messageId,
  ];
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

  // Apply alternating row color
  const bgColor = (newRow % 2 === 0) ? '#F9F9F9' : '#FFFFFF';
  sheet.getRange(newRow, 1, 1, rowData.length)
    .setBackground(bgColor)
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Format the date column
  sheet.getRange(newRow, 1).setNumberFormat('yyyy-mm-dd hh:mm');

  // Format the numeric amount column
  sheet.getRange(newRow, 5).setNumberFormat('$#,##0.00');

  // Apply status color to category cell
  applyCategoryStyle_(sheet, newRow, 6, item.category);
}

/**
 * Creates the branded invoice log sheet with styled headers and footer.
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @returns {Sheet} The created sheet.
 */
function createLogSheet_(ss) {
  const sheet = ss.insertSheet(LOG_SHEET_NAME);

  const headers = [
    'Date',
    'Sender',
    'Subject',
    'Amount',
    'Amount (Numeric)',
    'Category',
    'Matched Keyword',
    'Message ID',
  ];

  // Write headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Style header row (brand spec)
  headerRange
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Set row height for header
  sheet.setRowHeight(1, 32);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 150);  // Date
  sheet.setColumnWidth(2, 220);  // Sender
  sheet.setColumnWidth(3, 340);  // Subject
  sheet.setColumnWidth(4, 110);  // Amount
  sheet.setColumnWidth(5, 130);  // Amount (Numeric)
  sheet.setColumnWidth(6, 150);  // Category
  sheet.setColumnWidth(7, 150);  // Matched Keyword
  sheet.setColumnWidth(8, 160);  // Message ID

  // Hide the Message ID column (used internally for dedup display)
  sheet.hideColumns(8);

  // Add branded footer
  addBrandedFooter_(sheet, headers.length);

  // Move this sheet to position 1
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  return sheet;
}

/**
 * Adds the "Powered by TAKScripts" footer row.
 *
 * @param {Sheet} sheet - The target sheet.
 * @param {number} numCols - Number of columns to merge.
 */
function addBrandedFooter_(sheet, numCols) {
  const footerRow = 2; // Will be pushed down as data is added
  // We place a note in the sheet properties instead of a fixed row,
  // since rows shift. Use the sheet description or a named range.
  // For simplicity, we add it at row 1000 (far enough down).
  const range = sheet.getRange(1000, 1, 1, numCols);
  range.merge();
  range.setValue('Powered by TAKScripts · takscripts.store');
  range.setFontColor('#CCCCCC')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setFontFamily('Roboto');
}

/**
 * Applies a color-coded background to category cells.
 *
 * @param {Sheet} sheet - The log sheet.
 * @param {number} row - Row number.
 * @param {number} col - Column number.
 * @param {string} category - The invoice category.
 */
function applyCategoryStyle_(sheet, row, col, category) {
  const cell = sheet.getRange(row, col);
  const styles = {
    'Invoice':              { bg: '#FFF8E1', fg: '#F57F17' },   // Warning/amber
    'Receipt':              { bg: '#E8F5E9', fg: '#2E7D32' },   // Success/green
    'Payment Confirmation': { bg: '#E8F5E9', fg: '#2E7D32' },   // Success/green
    'Order Confirmation':   { bg: '#E3F2FD', fg: '#1565C0' },   // Info/blue
    'Subscription':         { bg: '#FFF8E1', fg: '#F57F17' },   // Warning/amber
    'Refund':               { bg: '#FFEBEE', fg: '#C62828' },   // Error/red
  };

  const style = styles[category] || { bg: '#F5F5F5', fg: '#666666' };
  cell.setBackground(style.bg)
    .setFontColor(style.fg)
    .setFontWeight('bold');
}


// ═══════════════════════════════════════════
// ALERT EMAILS
// ═══════════════════════════════════════════

/**
 * Sends a branded HTML alert email when an invoice exceeds the threshold.
 *
 * @param {Object} item - The detected invoice data.
 * @param {number} threshold - The configured alert threshold.
 * @param {string} alertEmail - Email address to send alert to.
 */
function sendThresholdAlert_(item, threshold, alertEmail) {
  if (!alertEmail) return;

  const subject = '⚠️ High Invoice Alert — ' + item.amount + ' from ' + item.senderEmail;

  const htmlBody = `
    <div style="font-family: 'Segoe UI', system-ui, sans-serif; max-width: 600px; margin: 0 auto;">
      <!-- Header -->
      <div style="background: #1A1A1A; padding: 24px; text-align: center; border-radius: 8px 8px 0 0;">
        <div style="font-size: 28px; margin-bottom: 4px;">🕷</div>
        <div style="color: #C9A84C; font-size: 16px; font-weight: 700; letter-spacing: 1px;">TAKScripts</div>
        <div style="color: #888; font-size: 11px; margin-top: 4px;">Email Invoice Detector · Alert</div>
      </div>

      <!-- Body -->
      <div style="background: #FFFFFF; border: 1px solid #eee; border-top: none; padding: 24px;">
        <h2 style="margin: 0 0 16px; font-size: 18px; color: #C62828;">
          ⚠️ Invoice Exceeds Threshold
        </h2>
        <p style="font-size: 14px; color: #333; line-height: 1.6; margin: 0 0 20px;">
          An email was detected with an amount of <strong>${item.amount}</strong>,
          which exceeds your alert threshold of <strong>$${threshold.toFixed(2)}</strong>.
        </p>

        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #888; width: 120px;">From</td>
            <td style="padding: 10px 0; color: #333;">${item.senderEmail}</td>
          </tr>
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #888;">Subject</td>
            <td style="padding: 10px 0; color: #333;">${escapeHtml_(item.subject)}</td>
          </tr>
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #888;">Amount</td>
            <td style="padding: 10px 0; color: #C62828; font-weight: 700; font-size: 16px;">${item.amount}</td>
          </tr>
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #888;">Category</td>
            <td style="padding: 10px 0; color: #333;">${item.category}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #888;">Date</td>
            <td style="padding: 10px 0; color: #333;">${item.date}</td>
          </tr>
        </table>
      </div>

      <!-- Footer -->
      <div style="background: #FAFAFA; border: 1px solid #eee; border-top: none; padding: 16px; text-align: center; border-radius: 0 0 8px 8px;">
        <p style="font-size: 11px; color: #999; margin: 0;">
          Sent by <strong>Email Invoice Detector</strong> via
          <a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none;">TAKScripts</a>
        </p>
      </div>
    </div>
  `;

  try {
    GmailApp.sendEmail(alertEmail, subject,
      'High Invoice Alert: ' + item.amount + ' from ' + item.senderEmail,
      { htmlBody: htmlBody }
    );
    Logger.log('Alert sent to ' + alertEmail + ' for amount ' + item.amount);
  } catch (e) {
    Logger.log('Failed to send alert email: ' + e.message);
  }
}


// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Extracts the email address from a "From" header field.
 * Handles formats like: "John Doe <john@example.com>" or "john@example.com"
 *
 * @param {string} fromField - The raw From header value.
 * @returns {string} The extracted email address.
 */
function extractEmail_(fromField) {
  const match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

/**
 * Formats a Date object to a human-readable string.
 *
 * @param {Date} date - The date to format.
 * @returns {string} Formatted date string (e.g., "March 15, 2026").
 */
function formatDate_(date) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

/**
 * Escapes HTML special characters to prevent XSS in alert emails.
 *
 * @param {string} str - The raw string.
 * @returns {string} The escaped string.
 */
function escapeHtml_(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Resets the processed message IDs. Use if you want to re-scan old emails.
 * Available from Script Editor only (not in the menu for safety).
 */
function resetProcessedIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
  Logger.log('Processed message IDs cleared. Next scan will re-process all emails.');
}


// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

/**
 * Returns the full HTML for the branded settings sidebar.
 * Follows TAKScripts design guide: dark header, gold accents,
 * white body, rounded inputs with gold focus states.
 *
 * @returns {string} Complete HTML string.
 */
function getSettingsHtml_() {
  const categoryOptions = Object.keys(DETECTION_RULES)
    .map(cat => `<label class="checkbox-row"><input type="checkbox" value="${cat}" checked /><span>${cat}</span></label>`)
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

    /* ── Header ────────────────────────── */
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

    /* ── Form ──────────────────────────── */
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
    .field .help {
      font-size: 11px;
      color: #999;
      margin-top: 4px;
      line-height: 1.4;
    }

    /* ── Inline fields ────────────────── */
    .inline-fields {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
    }

    /* ── Toggle ────────────────────────── */
    .toggle-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 0;
    }
    .toggle-label {
      font-size: 13px;
      color: #333;
    }
    .toggle-label .sublabel {
      display: block;
      font-size: 11px;
      color: #999;
      margin-top: 2px;
    }
    .toggle {
      position: relative;
      width: 40px;
      height: 22px;
      flex-shrink: 0;
      margin-left: 12px;
    }
    .toggle input {
      opacity: 0;
      width: 0;
      height: 0;
    }
    .toggle .slider {
      position: absolute;
      inset: 0;
      background: #ccc;
      border-radius: 22px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .toggle .slider::before {
      content: '';
      position: absolute;
      height: 16px;
      width: 16px;
      left: 3px;
      bottom: 3px;
      background: white;
      border-radius: 50%;
      transition: transform 0.2s;
    }
    .toggle input:checked + .slider {
      background: #C9A84C;
    }
    .toggle input:checked + .slider::before {
      transform: translateX(18px);
    }

    /* ── Checkboxes ────────────────────── */
    .checkbox-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 6px;
    }
    .checkbox-row {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #444;
      cursor: pointer;
      padding: 4px 0;
    }
    .checkbox-row input[type="checkbox"] {
      accent-color: #C9A84C;
      width: 14px;
      height: 14px;
    }

    /* ── Buttons ───────────────────────── */
    .btn {
      width: 100%;
      padding: 12px;
      border: none;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      letter-spacing: 0.5px;
    }
    .btn-primary {
      background: #1A1A1A;
      color: #C9A84C;
      border: 1px solid #C9A84C;
    }
    .btn-primary:hover {
      background: #C9A84C;
      color: #1A1A1A;
    }
    .btn-secondary {
      background: white;
      color: #666;
      border: 1px solid #ddd;
      margin-top: 8px;
    }
    .btn-secondary:hover {
      border-color: #999;
      color: #333;
    }

    /* ── Status ────────────────────────── */
    .status {
      text-align: center;
      padding: 8px;
      font-size: 12px;
      margin-top: 8px;
      border-radius: 6px;
      display: none;
    }
    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }
    .status.error   { display: block; background: #FFEBEE; color: #C62828; }

    .divider { border-top: 1px solid #eee; margin: 20px 0; }
  </style>
</head>
<body>
  <!-- ── Header ─────────────────────── -->
  <div class="header">
    <div class="logo">🕷</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">Email Invoice Detector · Settings</div>
  </div>

  <div class="form">

    <!-- ── Scan Settings ──────────────── -->
    <div class="section-title">Scan Settings</div>

    <div class="field">
      <label>Gmail Label (optional)</label>
      <input type="text" id="scanLabel" placeholder="Leave empty to scan Inbox" />
      <div class="help">Only scan emails with this label. Leave empty to scan your entire inbox.</div>
    </div>

    <div class="inline-fields">
      <div class="field">
        <label>Lookback (days)</label>
        <select id="scanAge">
          <option value="1">1 day</option>
          <option value="3">3 days</option>
          <option value="7" selected>7 days</option>
          <option value="14">14 days</option>
          <option value="30">30 days</option>
        </select>
      </div>
      <div class="field">
        <label>Scan Frequency</label>
        <select id="frequency">
          <option value="5">Every 5 min</option>
          <option value="15">Every 15 min</option>
          <option value="30">Every 30 min</option>
          <option value="60" selected>Every hour</option>
          <option value="360">Every 6 hours</option>
          <option value="720">Every 12 hours</option>
          <option value="1440">Once a day</option>
        </select>
      </div>
    </div>

    <div class="toggle-row">
      <div class="toggle-label">
        Auto-scan on schedule
        <span class="sublabel">Creates a time-driven trigger</span>
      </div>
      <label class="toggle">
        <input type="checkbox" id="enableAutoTrigger" />
        <span class="slider"></span>
      </label>
    </div>

    <!-- ── Categories ─────────────────── -->
    <div class="section-title">Detection Categories</div>

    <div class="checkbox-grid" id="categoryGrid">
      ${categoryOptions}
    </div>

    <!-- ── Alerts ─────────────────────── -->
    <div class="section-title">Amount Alerts</div>

    <div class="toggle-row">
      <div class="toggle-label">
        High-amount email alerts
        <span class="sublabel">Get notified for large invoices</span>
      </div>
      <label class="toggle">
        <input type="checkbox" id="enableAlerts" />
        <span class="slider"></span>
      </label>
    </div>

    <div class="inline-fields">
      <div class="field">
        <label>Threshold ($)</label>
        <input type="number" id="alertThreshold" placeholder="500" min="0" step="50" />
      </div>
      <div class="field">
        <label>Alert Email</label>
        <input type="email" id="alertEmail" placeholder="you@example.com" />
      </div>
    </div>

    <!-- ── Actions ─────────────────────── -->
    <div class="divider"></div>

    <button class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    /**
     * Load saved settings on sidebar open.
     */
    google.script.run.withSuccessHandler(function(s) {
      document.getElementById('scanLabel').value = s.scanLabel || '';
      document.getElementById('scanAge').value = s.scanAge || '7';
      document.getElementById('frequency').value = s.frequency || '60';
      document.getElementById('enableAutoTrigger').checked = s.enableAutoTrigger === true || s.enableAutoTrigger === 'true';
      document.getElementById('enableAlerts').checked = s.enableAlerts === true || s.enableAlerts === 'true';
      document.getElementById('alertThreshold').value = s.alertThreshold || '500';
      document.getElementById('alertEmail').value = s.alertEmail || '';

      // Restore category checkboxes
      if (s.categories) {
        var enabled = s.categories.split(',').map(function(c) { return c.trim(); });
        var boxes = document.querySelectorAll('#categoryGrid input[type="checkbox"]');
        boxes.forEach(function(box) {
          box.checked = enabled.indexOf(box.value) !== -1;
        });
      }
    }).loadSettings();

    /**
     * Collect form values and save via server call.
     */
    function save() {
      // Gather enabled categories
      var boxes = document.querySelectorAll('#categoryGrid input[type="checkbox"]');
      var cats = [];
      boxes.forEach(function(box) { if (box.checked) cats.push(box.value); });

      var settings = {
        scanLabel: document.getElementById('scanLabel').value.trim(),
        scanAge: document.getElementById('scanAge').value,
        frequency: document.getElementById('frequency').value,
        enableAutoTrigger: document.getElementById('enableAutoTrigger').checked,
        enableAlerts: document.getElementById('enableAlerts').checked,
        alertThreshold: document.getElementById('alertThreshold').value,
        alertEmail: document.getElementById('alertEmail').value.trim(),
        categories: cats.join(', '),
      };

      var statusEl = document.getElementById('status');
      statusEl.className = 'status';
      statusEl.style.display = 'none';

      google.script.run
        .withSuccessHandler(function() {
          statusEl.textContent = '✓ Settings saved successfully';
          statusEl.className = 'status success';
        })
        .withFailureHandler(function(err) {
          statusEl.textContent = '✕ Error: ' + err.message;
          statusEl.className = 'status error';
        })
        .saveSettings(settings);
    }
  </script>
</body>
</html>`;
}
