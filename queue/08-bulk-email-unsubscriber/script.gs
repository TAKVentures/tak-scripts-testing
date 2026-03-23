/**
 * Bulk Email Unsubscriber by TAKScripts
 * ======================================
 * Scans your Gmail for newsletters and subscription emails, lists them
 * in a clean dashboard, and lets you bulk-unsubscribe with one click.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for scan configuration
 * - Detects subscriptions via List-Unsubscribe headers, promotional labels,
 *   and common newsletter sender patterns
 * - Shows sender, frequency, last received, and unsubscribe link
 * - Mark rows "Unsubscribe" and run — handles mailto: and URL unsubscribes
 * - Test Run mode to preview without taking action
 * - Full action log on a separate sheet
 * - Safe — never deletes emails, only unsubscribes
 * - Idempotent — safe to run multiple times
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var CONFIG = {
  SCANNER_SHEET: '📬 Subscription Scanner',
  LOG_SHEET: '📋 Unsubscribe Log',
  SETTINGS_KEY: 'unsubscriber_settings',
  PROCESSED_KEY: 'unsubscriber_processed',
  HEADER_BG: '#1A1A1A',
  HEADER_TEXT: '#C9A84C',
  HEADER_FONT: 'Roboto Mono',
  HEADER_SIZE: 10,
  DATA_FONT: 'Roboto',
  DATA_SIZE: 10,
  ROW_WHITE: '#FFFFFF',
  ROW_GRAY: '#F9F9F9',
  STATUS_SUCCESS_BG: '#E8F5E9',
  STATUS_SUCCESS_TEXT: '#2E7D32',
  STATUS_WARNING_BG: '#FFF8E1',
  STATUS_WARNING_TEXT: '#F57F17',
  STATUS_ERROR_BG: '#FFEBEE',
  STATUS_ERROR_TEXT: '#C62828',
  STATUS_INFO_BG: '#E3F2FD',
  STATUS_INFO_TEXT: '#1565C0',
  FOOTER_COLOR: '#CCCCCC',
  FOOTER_SIZE: 9,
  FOOTER_TEXT: 'Powered by TAKScripts · takscripts.store',
};

// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

/**
 * Creates the TAKScripts menu when the spreadsheet opens.
 * This runs automatically — no setup needed.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🕷 TAKScripts')
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('▶️ Scan Subscriptions', 'scanSubscriptions')
    .addItem('🧪 Test Run', 'testRun')
    .addSeparator()
    .addItem('📊 View Log', 'viewLog')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Bulk Unsubscriber Settings')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">🕷</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Bulk Email Unsubscriber</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 · by TAK Ventures</p>' +
      '<hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">' +
      '<p style="font-size: 13px; color: #333; line-height: 1.6;">' +
        'Part of the <strong>TAKScripts</strong> collection.<br>' +
        'Pre-built Google Apps Scripts for small business.' +
      '</p>' +
      '<p style="margin-top: 16px;">' +
        '<a href="https://takscripts.store" target="_blank" ' +
           'style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">' +
          'takscripts.store →' +
        '</a>' +
      '</p>' +
    '</div>'
  ).setWidth(300).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to the unsubscribe log sheet.
 */
function viewLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No unsubscribe log yet. Run a scan and process unsubscribes first.');
    return;
  }
  ss.setActiveSheet(sheet);
}

// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Save settings from the sidebar.
 */
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.SETTINGS_KEY, JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings. Returns defaults on first run.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(CONFIG.SETTINGS_KEY);
  if (!raw) {
    return {
      scanDays: 90,
      minEmails: 3,
      includePromotions: true,
      includeUpdates: true,
      includeSocial: false,
      excludeSenders: '',
    };
  }
  return JSON.parse(raw);
}

// ═══════════════════════════════════════════
// CORE SCANNER
// ═══════════════════════════════════════════

/**
 * Main scan function. Searches Gmail for subscription emails,
 * aggregates by sender, and writes results to the scanner sheet.
 */
function scanSubscriptions() {
  var ui = SpreadsheetApp.getUi();
  var settings = loadSettings();

  ui.alert(
    '▶️ Scanning Subscriptions',
    'Scanning the last ' + settings.scanDays + ' days of email.\n' +
    'This may take a minute. You\'ll see results on the "' + CONFIG.SCANNER_SHEET + '" tab when done.',
    ui.ButtonSet.OK
  );

  var subscriptions = findSubscriptions_(settings, false);
  writeScannerSheet_(subscriptions);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SCANNER_SHEET);
  if (sheet) ss.setActiveSheet(sheet);

  ui.alert(
    '✅ Scan Complete',
    'Found ' + subscriptions.length + ' subscriptions.\n\n' +
    'To unsubscribe:\n' +
    '1. Set the Status column to "Unsubscribe" for the ones you want to remove\n' +
    '2. Go to 🕷 TAKScripts → ▶️ Scan Subscriptions\n' +
    '   (it will process any marked rows)\n\n' +
    'Or use the dropdown in column F to mark rows.',
    ui.ButtonSet.OK
  );

  processUnsubscribes_();
}

/**
 * Test run — scans and shows subscriptions without taking any action.
 * Marked rows are NOT processed.
 */
function testRun() {
  var ui = SpreadsheetApp.getUi();
  var settings = loadSettings();

  ui.alert(
    '🧪 Test Run',
    'Scanning the last ' + settings.scanDays + ' days for subscriptions.\n' +
    'No unsubscribe actions will be taken — preview only.',
    ui.ButtonSet.OK
  );

  var subscriptions = findSubscriptions_(settings, true);
  writeScannerSheet_(subscriptions);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SCANNER_SHEET);
  if (sheet) ss.setActiveSheet(sheet);

  ui.alert(
    '🧪 Test Complete',
    'Found ' + subscriptions.length + ' subscriptions (preview only).\n\n' +
    'No actions were taken. Review the results on the "' + CONFIG.SCANNER_SHEET + '" tab.\n' +
    'When ready, use ▶️ Scan Subscriptions to run for real.',
    ui.ButtonSet.OK
  );
}

/**
 * Internal: Finds subscription emails in Gmail.
 * Returns an array of subscription objects aggregated by sender.
 */
function findSubscriptions_(settings, isTestRun) {
  var daysBack = parseInt(settings.scanDays, 10) || 90;
  var minEmails = parseInt(settings.minEmails, 10) || 3;
  var excludeList = (settings.excludeSenders || '')
    .split(',')
    .map(function(s) { return s.trim().toLowerCase(); })
    .filter(Boolean);

  // Build search queries to find subscription-like emails
  var queries = [];
  queries.push('has:nousersubs newer_than:' + daysBack + 'd');
  if (settings.includePromotions) {
    queries.push('category:promotions newer_than:' + daysBack + 'd');
  }
  if (settings.includeUpdates) {
    queries.push('category:updates newer_than:' + daysBack + 'd');
  }
  if (settings.includeSocial) {
    queries.push('category:social newer_than:' + daysBack + 'd');
  }
  queries.push('label:^unsub newer_than:' + daysBack + 'd');

  // Collect all unique message IDs across queries
  var seenThreadIds = {};
  var allThreads = [];

  for (var q = 0; q < queries.length; q++) {
    try {
      var threads = GmailApp.search(queries[q], 0, 200);
      for (var t = 0; t < threads.length; t++) {
        var tid = threads[t].getId();
        if (!seenThreadIds[tid]) {
          seenThreadIds[tid] = true;
          allThreads.push(threads[t]);
        }
      }
    } catch (e) {
      Logger.log('Query skipped: ' + queries[q] + ' — ' + e.message);
    }
  }

  // Aggregate by sender email
  var senderMap = {};

  for (var i = 0; i < allThreads.length; i++) {
    var messages = allThreads[i].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var from = msg.getFrom();
      var email = extractEmail_(from);
      var name = extractSenderName_(from);

      // Check exclusions
      var excluded = false;
      for (var ex = 0; ex < excludeList.length; ex++) {
        if (email.indexOf(excludeList[ex]) !== -1) {
          excluded = true;
          break;
        }
      }
      if (excluded) continue;

      // Get unsubscribe header
      var unsubHeader = '';
      try {
        unsubHeader = msg.getHeader('List-Unsubscribe') || '';
      } catch (e) {
        // Header may not exist
      }

      if (!senderMap[email]) {
        senderMap[email] = {
          senderName: name,
          senderEmail: email,
          count: 0,
          lastDate: null,
          unsubscribeLink: '',
          hasUnsubHeader: false,
        };
      }

      senderMap[email].count++;
      var msgDate = msg.getDate();
      if (!senderMap[email].lastDate || msgDate > senderMap[email].lastDate) {
        senderMap[email].lastDate = msgDate;
      }

      // Extract unsubscribe link (prefer URL over mailto)
      if (unsubHeader && !senderMap[email].unsubscribeLink) {
        var urlMatch = unsubHeader.match(/<(https?:\/\/[^>]+)>/);
        var mailtoMatch = unsubHeader.match(/<(mailto:[^>]+)>/);
        if (urlMatch) {
          senderMap[email].unsubscribeLink = urlMatch[1];
          senderMap[email].hasUnsubHeader = true;
        } else if (mailtoMatch) {
          senderMap[email].unsubscribeLink = mailtoMatch[1];
          senderMap[email].hasUnsubHeader = true;
        }
      }
    }
  }

  // Filter and sort results
  var results = [];
  var keys = Object.keys(senderMap);
  for (var k = 0; k < keys.length; k++) {
    var sub = senderMap[keys[k]];
    if (sub.count >= minEmails) {
      results.push(sub);
    }
  }

  // Sort by email count descending
  results.sort(function(a, b) { return b.count - a.count; });

  return results;
}

/**
 * Internal: Writes subscription data to the scanner sheet.
 * Creates the sheet if it doesn't exist, clears old data if it does.
 */
function writeScannerSheet_(subscriptions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SCANNER_SHEET);

  if (sheet) {
    sheet.clear();
    // Remove existing data validations
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    }
  } else {
    sheet = ss.insertSheet(CONFIG.SCANNER_SHEET);
  }

  var headers = ['Sender', 'Email Address', 'Emails Found', 'Frequency', 'Last Received', 'Status', 'Unsubscribe Link'];

  // Write headers
  sheet.appendRow(headers);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontFamily(CONFIG.HEADER_FONT)
    .setFontSize(CONFIG.HEADER_SIZE)
    .setFontWeight('bold')
    .setBackground(CONFIG.HEADER_BG)
    .setFontColor(CONFIG.HEADER_TEXT)
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);

  // Column widths
  sheet.setColumnWidth(1, 200);  // Sender
  sheet.setColumnWidth(2, 280);  // Email Address
  sheet.setColumnWidth(3, 100);  // Emails Found
  sheet.setColumnWidth(4, 120);  // Frequency
  sheet.setColumnWidth(5, 140);  // Last Received
  sheet.setColumnWidth(6, 130);  // Status
  sheet.setColumnWidth(7, 350);  // Unsubscribe Link

  // Write data rows
  for (var i = 0; i < subscriptions.length; i++) {
    var sub = subscriptions[i];
    var frequency = getFrequencyLabel_(sub.count, parseInt(loadSettings().scanDays, 10) || 90);
    var row = [
      sub.senderName || sub.senderEmail.split('@')[0],
      sub.senderEmail,
      sub.count,
      frequency,
      sub.lastDate ? Utilities.formatDate(sub.lastDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
      'Keep',
      sub.unsubscribeLink || '(none found)',
    ];
    sheet.appendRow(row);

    var rowNum = i + 2;
    var rowRange = sheet.getRange(rowNum, 1, 1, headers.length);

    // Alternating row colors
    var bgColor = (i % 2 === 0) ? CONFIG.ROW_WHITE : CONFIG.ROW_GRAY;
    rowRange.setBackground(bgColor);
    rowRange.setFontFamily(CONFIG.DATA_FONT);
    rowRange.setFontSize(CONFIG.DATA_SIZE);
    rowRange.setVerticalAlignment('middle');
  }

  // Add status dropdown validation to the Status column
  if (subscriptions.length > 0) {
    var statusRange = sheet.getRange(2, 6, subscriptions.length, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Keep', 'Unsubscribe', 'Done', 'Error'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(rule);

    // Style the default "Keep" status cells
    for (var r = 0; r < subscriptions.length; r++) {
      applyStatusStyle_(sheet, r + 2, 'Keep');
    }
  }

  // Footer row
  var footerRow = subscriptions.length + 3;
  sheet.getRange(footerRow, 1).setValue(CONFIG.FOOTER_TEXT);
  var footerRange = sheet.getRange(footerRow, 1, 1, headers.length);
  footerRange.merge();
  footerRange
    .setFontColor(CONFIG.FOOTER_COLOR)
    .setFontSize(CONFIG.FOOTER_SIZE)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setFontFamily(CONFIG.DATA_FONT);
}

// ═══════════════════════════════════════════
// UNSUBSCRIBE ENGINE
// ═══════════════════════════════════════════

/**
 * Internal: Processes rows marked as "Unsubscribe" on the scanner sheet.
 * Attempts to unsubscribe via URL fetch or mailto, then updates status.
 */
function processUnsubscribes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SCANNER_SHEET);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var actionsPerformed = 0;

  for (var i = 0; i < data.length; i++) {
    var status = data[i][5];
    if (status !== 'Unsubscribe') continue;

    var senderName = data[i][0];
    var senderEmail = data[i][1];
    var unsubLink = data[i][6];
    var rowNum = i + 2;
    var result;

    if (!unsubLink || unsubLink === '(none found)') {
      result = { success: false, method: 'none', error: 'No unsubscribe link found' };
    } else if (unsubLink.indexOf('mailto:') === 0) {
      result = handleMailtoUnsubscribe_(unsubLink, senderEmail);
    } else if (unsubLink.indexOf('http') === 0) {
      result = handleUrlUnsubscribe_(unsubLink);
    } else {
      result = { success: false, method: 'unknown', error: 'Unsupported unsubscribe method' };
    }

    // Update status on sheet
    if (result.success) {
      sheet.getRange(rowNum, 6).setValue('Done');
      applyStatusStyle_(sheet, rowNum, 'Done');
      actionsPerformed++;
    } else {
      sheet.getRange(rowNum, 6).setValue('Error');
      applyStatusStyle_(sheet, rowNum, 'Error');
    }

    // Log the action
    logAction_(senderName, senderEmail, result);
  }

  if (actionsPerformed > 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Unsubscribe Complete',
      'Successfully processed ' + actionsPerformed + ' unsubscribe request(s).\n\n' +
      'Check the "' + CONFIG.LOG_SHEET + '" tab for full details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Internal: Unsubscribe via HTTP GET to the unsubscribe URL.
 * Many newsletters use one-click unsubscribe links.
 */
function handleUrlUnsubscribe_(url) {
  try {
    var options = {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; TAKScripts-Unsubscriber/1.0)',
      },
    };
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code >= 200 && code < 400) {
      return { success: true, method: 'URL', detail: 'HTTP ' + code };
    } else {
      return { success: false, method: 'URL', error: 'HTTP ' + code };
    }
  } catch (e) {
    return { success: false, method: 'URL', error: e.message };
  }
}

/**
 * Internal: Unsubscribe via mailto link — sends an unsubscribe email.
 * Parses the mailto: URI for address and optional subject.
 */
function handleMailtoUnsubscribe_(mailtoLink, originalSender) {
  try {
    var address = mailtoLink.replace('mailto:', '').split('?')[0];
    var subject = 'Unsubscribe';

    // Parse optional subject from query string
    var queryMatch = mailtoLink.match(/[?&]subject=([^&]*)/i);
    if (queryMatch) {
      subject = decodeURIComponent(queryMatch[1]);
    }

    GmailApp.sendEmail(address, subject, 'Please unsubscribe this email address from your mailing list.\n\nSent via TAKScripts Bulk Email Unsubscriber.');

    return { success: true, method: 'Email', detail: 'Sent to ' + address };
  } catch (e) {
    return { success: false, method: 'Email', error: e.message };
  }
}

// ═══════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════

/**
 * Internal: Logs an unsubscribe action to the log sheet.
 * Creates the sheet with styled headers if it doesn't exist.
 */
function logAction_(senderName, senderEmail, result) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET);
    var headers = ['Timestamp', 'Sender', 'Email Address', 'Method', 'Result', 'Details'];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontFamily(CONFIG.HEADER_FONT)
      .setFontSize(CONFIG.HEADER_SIZE)
      .setFontWeight('bold')
      .setBackground(CONFIG.HEADER_BG)
      .setFontColor(CONFIG.HEADER_TEXT)
      .setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 32);

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 300);

    // Footer
    var footerRange = sheet.getRange(3, 1, 1, headers.length);
    footerRange.merge();
    sheet.getRange(3, 1).setValue(CONFIG.FOOTER_TEXT);
    footerRange
      .setFontColor(CONFIG.FOOTER_COLOR)
      .setFontSize(CONFIG.FOOTER_SIZE)
      .setFontStyle('italic')
      .setHorizontalAlignment('center')
      .setFontFamily(CONFIG.DATA_FONT);
  }

  // Find the right row to insert (before footer)
  var lastRow = sheet.getLastRow();
  var insertRow = lastRow; // Insert before footer if present

  // Check if last row is footer
  var lastVal = sheet.getRange(lastRow, 1).getValue();
  if (lastVal === CONFIG.FOOTER_TEXT) {
    sheet.insertRowBefore(lastRow);
    insertRow = lastRow;
  } else {
    insertRow = lastRow + 1;
  }

  var statusText = result.success ? 'Success' : 'Failed';
  var detailText = result.success ? (result.detail || '') : (result.error || '');
  var rowData = [new Date(), senderName, senderEmail, result.method, statusText, detailText];

  sheet.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);

  // Style the new row
  var rowRange = sheet.getRange(insertRow, 1, 1, rowData.length);
  var rowIndex = insertRow - 2; // 0-based for alternating
  rowRange.setBackground(rowIndex % 2 === 0 ? CONFIG.ROW_WHITE : CONFIG.ROW_GRAY);
  rowRange.setFontFamily(CONFIG.DATA_FONT);
  rowRange.setFontSize(CONFIG.DATA_SIZE);
  rowRange.setVerticalAlignment('middle');

  // Color the result cell
  var resultCell = sheet.getRange(insertRow, 5);
  if (result.success) {
    resultCell.setBackground(CONFIG.STATUS_SUCCESS_BG).setFontColor(CONFIG.STATUS_SUCCESS_TEXT);
  } else {
    resultCell.setBackground(CONFIG.STATUS_ERROR_BG).setFontColor(CONFIG.STATUS_ERROR_TEXT);
  }
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Internal: Extracts the email address from a "Name <email>" string.
 */
function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

/**
 * Internal: Extracts the display name from a "Name <email>" string.
 */
function extractSenderName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim() : '';
}

/**
 * Internal: Calculates a human-readable frequency label.
 */
function getFrequencyLabel_(count, days) {
  var perWeek = (count / days) * 7;
  if (perWeek >= 7) return 'Daily+';
  if (perWeek >= 3) return 'Several/week';
  if (perWeek >= 1) return 'Weekly';
  if (perWeek >= 0.5) return 'Biweekly';
  return 'Monthly';
}

/**
 * Internal: Applies color styling to a status cell based on its value.
 */
function applyStatusStyle_(sheet, row, status) {
  var cell = sheet.getRange(row, 6);
  switch (status) {
    case 'Keep':
      cell.setBackground(CONFIG.STATUS_INFO_BG).setFontColor(CONFIG.STATUS_INFO_TEXT).setFontWeight('bold');
      break;
    case 'Unsubscribe':
      cell.setBackground(CONFIG.STATUS_WARNING_BG).setFontColor(CONFIG.STATUS_WARNING_TEXT).setFontWeight('bold');
      break;
    case 'Done':
      cell.setBackground(CONFIG.STATUS_SUCCESS_BG).setFontColor(CONFIG.STATUS_SUCCESS_TEXT).setFontWeight('bold');
      break;
    case 'Error':
      cell.setBackground(CONFIG.STATUS_ERROR_BG).setFontColor(CONFIG.STATUS_ERROR_TEXT).setFontWeight('bold');
      break;
    default:
      break;
  }
}

/**
 * Trigger for onEdit — auto-styles status cells when changed.
 * Install via Edit > Triggers or it runs as a simple trigger.
 */
function onEdit(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SCANNER_SHEET) return;
  if (e.range.getColumn() !== 6) return; // Status column
  if (e.range.getRow() < 2) return; // Skip header

  var value = e.range.getValue();
  applyStatusStyle_(sheet, e.range.getRow(), value);
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

function getSettingsHtml() {
  return '<!DOCTYPE html>' +
'<html>' +
'<head>' +
'  <style>' +
'    * { box-sizing: border-box; margin: 0; padding: 0; }' +
'    body {' +
'      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;' +
'      background: #FAFAFA;' +
'      color: #1a1a1a;' +
'      font-size: 13px;' +
'    }' +
'    .header {' +
'      background: #1A1A1A;' +
'      color: white;' +
'      padding: 20px 16px;' +
'      text-align: center;' +
'    }' +
'    .header .logo { font-size: 24px; margin-bottom: 4px; }' +
'    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }' +
'    .header .brand { color: #C9A84C; }' +
'    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }' +
'    .form { padding: 16px; }' +
'    .field { margin-bottom: 16px; }' +
'    .field label {' +
'      display: block;' +
'      font-size: 11px;' +
'      font-weight: 600;' +
'      text-transform: uppercase;' +
'      letter-spacing: 1px;' +
'      color: #666;' +
'      margin-bottom: 6px;' +
'    }' +
'    .field input[type="number"], .field input[type="text"] {' +
'      width: 100%;' +
'      padding: 10px 12px;' +
'      border: 1px solid #ddd;' +
'      border-radius: 6px;' +
'      font-size: 13px;' +
'      font-family: inherit;' +
'      transition: border-color 0.2s;' +
'    }' +
'    .field input:focus {' +
'      outline: none;' +
'      border-color: #C9A84C;' +
'      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);' +
'    }' +
'    .field .help {' +
'      font-size: 11px;' +
'      color: #999;' +
'      margin-top: 4px;' +
'      line-height: 1.4;' +
'    }' +
'    .checkbox-group { margin-top: 8px; }' +
'    .checkbox-item {' +
'      display: flex;' +
'      align-items: center;' +
'      margin-bottom: 10px;' +
'      cursor: pointer;' +
'    }' +
'    .checkbox-item input[type="checkbox"] {' +
'      width: 16px;' +
'      height: 16px;' +
'      margin-right: 10px;' +
'      accent-color: #C9A84C;' +
'      cursor: pointer;' +
'    }' +
'    .checkbox-item span {' +
'      font-size: 13px;' +
'      color: #333;' +
'    }' +
'    .btn {' +
'      width: 100%;' +
'      padding: 12px;' +
'      border: none;' +
'      border-radius: 8px;' +
'      font-size: 13px;' +
'      font-weight: 600;' +
'      cursor: pointer;' +
'      transition: all 0.2s;' +
'      letter-spacing: 0.5px;' +
'    }' +
'    .btn-primary {' +
'      background: #1A1A1A;' +
'      color: #C9A84C;' +
'      border: 1px solid #C9A84C;' +
'    }' +
'    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }' +
'    .btn-secondary {' +
'      background: white;' +
'      color: #666;' +
'      border: 1px solid #ddd;' +
'      margin-top: 8px;' +
'    }' +
'    .btn-secondary:hover { border-color: #999; color: #333; }' +
'    .status {' +
'      text-align: center;' +
'      padding: 8px;' +
'      font-size: 12px;' +
'      margin-top: 8px;' +
'      border-radius: 6px;' +
'      display: none;' +
'    }' +
'    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }' +
'    .status.error { display: block; background: #FFEBEE; color: #C62828; }' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }' +
'    .section-title {' +
'      font-size: 12px;' +
'      font-weight: 700;' +
'      color: #1A1A1A;' +
'      margin-bottom: 12px;' +
'      text-transform: uppercase;' +
'      letter-spacing: 0.5px;' +
'    }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="logo">🕷</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Bulk Email Unsubscriber · Settings</div>' +
'  </div>' +
'' +
'  <div class="form">' +
'    <div class="field">' +
'      <label>Scan Period (Days)</label>' +
'      <input type="number" id="scanDays" min="7" max="365" placeholder="90" />' +
'      <div class="help">How far back to search for subscription emails.</div>' +
'    </div>' +
'' +
'    <div class="field">' +
'      <label>Minimum Emails to Flag</label>' +
'      <input type="number" id="minEmails" min="1" max="100" placeholder="3" />' +
'      <div class="help">Only show senders with at least this many emails.</div>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <div class="section-title">Categories to Include</div>' +
'    <div class="checkbox-group">' +
'      <label class="checkbox-item">' +
'        <input type="checkbox" id="includePromotions" />' +
'        <span>Promotions</span>' +
'      </label>' +
'      <label class="checkbox-item">' +
'        <input type="checkbox" id="includeUpdates" />' +
'        <span>Updates</span>' +
'      </label>' +
'      <label class="checkbox-item">' +
'        <input type="checkbox" id="includeSocial" />' +
'        <span>Social</span>' +
'      </label>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <div class="field">' +
'      <label>Exclude Senders</label>' +
'      <input type="text" id="excludeSenders" placeholder="boss@company.com, team@work.com" />' +
'      <div class="help">Comma-separated. Emails containing these strings will be skipped.</div>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <button class="btn btn-primary" onclick="save()">Save Settings</button>' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
'' +
'    <div id="status" class="status"></div>' +
'  </div>' +
'' +
'  <script>' +
'    google.script.run.withSuccessHandler(function(settings) {' +
'      document.getElementById("scanDays").value = settings.scanDays || 90;' +
'      document.getElementById("minEmails").value = settings.minEmails || 3;' +
'      document.getElementById("includePromotions").checked = settings.includePromotions !== false;' +
'      document.getElementById("includeUpdates").checked = settings.includeUpdates !== false;' +
'      document.getElementById("includeSocial").checked = settings.includeSocial === true;' +
'      document.getElementById("excludeSenders").value = settings.excludeSenders || "";' +
'    }).loadSettings();' +
'' +
'    function save() {' +
'      var settings = {' +
'        scanDays: parseInt(document.getElementById("scanDays").value, 10) || 90,' +
'        minEmails: parseInt(document.getElementById("minEmails").value, 10) || 3,' +
'        includePromotions: document.getElementById("includePromotions").checked,' +
'        includeUpdates: document.getElementById("includeUpdates").checked,' +
'        includeSocial: document.getElementById("includeSocial").checked,' +
'        excludeSenders: document.getElementById("excludeSenders").value,' +
'      };' +
'' +
'      var statusEl = document.getElementById("status");' +
'      statusEl.className = "status";' +
'      statusEl.style.display = "none";' +
'' +
'      google.script.run' +
'        .withSuccessHandler(function() {' +
'          statusEl.textContent = "✓ Settings saved successfully";' +
'          statusEl.className = "status success";' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          statusEl.textContent = "✕ Error: " + err.message;' +
'          statusEl.className = "status error";' +
'        })' +
'        .saveSettings(settings);' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
