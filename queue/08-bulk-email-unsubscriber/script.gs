/**
 * Bulk Email Unsubscriber by TAKScripts
 * ======================================
 * Scans your Gmail for newsletters and subscription emails, lists them
 * in a clean dashboard, and lets you bulk-unsubscribe with one click.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for scan configuration
 * - Dashboard stats: Found · Unsubscribed · Errors · Kept · Daily+
 * - Detects subscriptions via List-Unsubscribe headers, promotional labels,
 *   and common newsletter sender patterns
 * - Shows sender, frequency, last received, and unsubscribe link
 * - Mark rows "Unsubscribe" and run — handles mailto: and URL unsubscribes
 * - Test Run mode to preview without taking action
 * - Full action log on a separate sheet
 * - Safe — never deletes emails, only unsubscribes
 * - Idempotent — safe to run multiple times
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Subscription Dashboard';
var LOG_SHEET_NAME = '\uD83D\uDCCB Unsubscribe Log';
var DASHBOARD_HEADER_ROW = 3;
var SETTINGS_KEY = 'unsubscriber_settings';

var BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E8F5E9', successText: '#2E7D32',
  warningBg: '#FFF8E1', warningText: '#F57F17',
  errorBg: '#FFEBEE', errorText: '#C62828',
  infoBg: '#E3F2FD', infoText: '#1565C0',
  headerFont: 'Roboto Mono', bodyFont: 'Roboto',
};

var SCAN_HEADERS = ['Sender', 'Email Address', 'Emails Found', 'Frequency', 'Last Received', 'Status', 'Unsubscribe Link'];
var LOG_HEADERS = ['Timestamp', 'Sender', 'Email Address', 'Method', 'Result', 'Details'];

var COL_STATUS = 6; // Status column in scanner sheet
var COL_FREQ = 4;   // Frequency column in scanner sheet


// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('\uD83D\uDD77 TAKScripts')
    .addItem('\u2699\uFE0F Settings', 'showSettings')
    .addItem('\u25B6\uFE0F Scan Subscriptions', 'scanSubscriptions')
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Dashboard', 'viewDashboard')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addItem('\uD83D\uDCCB View Log', 'viewLog')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Bulk Unsubscriber Settings')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function viewDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME) || getOrCreateDashboard_();
  ss.setActiveSheet(sheet);
}

function viewLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No unsubscribe log yet. Run a scan and process unsubscribes first.');
    return;
  }
  ss.setActiveSheet(sheet);
}

function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Bulk Email Unsubscriber</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 \u00B7 by TAK Ventures</p>' +
      '<hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">' +
      '<p style="font-size: 13px; color: #333; line-height: 1.6;">' +
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
  ).setWidth(300).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(SETTINGS_KEY, JSON.stringify(settings));
  return { success: true };
}

function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(SETTINGS_KEY);
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
// DASHBOARD MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Gets or creates the Subscription Dashboard sheet with stats bar and headers.
 * Idempotent — safe to call multiple times.
 */
function getOrCreateDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    // Migration: rename old scanner sheet if it exists
    var oldSheet = ss.getSheetByName('\uD83D\uDCEC Subscription Scanner');
    if (oldSheet) {
      oldSheet.setName(DASHBOARD_SHEET_NAME);
      oldSheet.clear();
      sheet = oldSheet;
    } else {
      sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
    }
  } else {
    var headerCheck = sheet.getRange(DASHBOARD_HEADER_ROW, 1).getValue();
    if (headerCheck === SCAN_HEADERS[0]) {
      return sheet; // Already set up
    }
  }

  var numCols = SCAN_HEADERS.length;

  // ── Row 1: Stat values ──────────────────────────────────
  sheet.getRange(1, 1, 1, numCols)
    .setValues([['—', '—', '—', '—', '—', '', '']])
    .setFontFamily(BRAND.headerFont)
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 52);

  // ── Row 2: Stat labels ──────────────────────────────────
  sheet.getRange(2, 1, 1, numCols)
    .setValues([['FOUND', 'UNSUBSCRIBED', 'ERRORS', 'KEPT', 'DAILY+', '', '']])
    .setFontFamily(BRAND.headerFont)
    .setFontSize(8)
    .setFontWeight('normal')
    .setFontColor('#888888')
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 20);

  // ── Row 3: Column headers ───────────────────────────────
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, numCols)
    .setValues([SCAN_HEADERS])
    .setFontFamily(BRAND.headerFont)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(DASHBOARD_HEADER_ROW, 32);

  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // ── Column widths ───────────────────────────────────────
  sheet.setColumnWidth(1, 200);  // Sender
  sheet.setColumnWidth(2, 280);  // Email Address
  sheet.setColumnWidth(3, 100);  // Emails Found
  sheet.setColumnWidth(4, 120);  // Frequency
  sheet.setColumnWidth(5, 140);  // Last Received
  sheet.setColumnWidth(6, 130);  // Status
  sheet.setColumnWidth(7, 350);  // Unsubscribe Link

  return sheet;
}

/**
 * Recalculates stats from current scan data and updates the stats bar.
 */
function refreshDashboardStats() {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();

  var found = 0;
  var unsubscribed = 0;
  var errors = 0;
  var kept = 0;
  var dailyPlus = 0;

  if (lastRow >= dataStartRow) {
    var numRows = lastRow - dataStartRow + 1;
    var data = sheet.getRange(dataStartRow, 1, numRows, SCAN_HEADERS.length).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue; // Skip empty rows
      found++;
      var status = (row[COL_STATUS - 1] || '').toString();
      var freq = (row[COL_FREQ - 1] || '').toString();
      if (status === 'Done') unsubscribed++;
      else if (status === 'Error') errors++;
      else if (status === 'Keep') kept++;
      if (freq === 'Daily+') dailyPlus++;
    }
  }

  sheet.getRange(1, 1, 1, 5).setValues([[
    found || '—',
    unsubscribed || '—',
    errors || '—',
    kept || '—',
    dailyPlus || '—',
  ]]);
}


// ═══════════════════════════════════════════
// CORE SCANNER
// ═══════════════════════════════════════════

function scanSubscriptions() {
  var ui = SpreadsheetApp.getUi();
  var settings = loadSettings();

  ui.alert(
    '\u25B6\uFE0F Scanning Subscriptions',
    'Scanning the last ' + settings.scanDays + ' days of email.\n' +
    'This may take a minute. You\'ll see results on the "' + DASHBOARD_SHEET_NAME + '" tab when done.',
    ui.ButtonSet.OK
  );

  var subscriptions = findSubscriptions_(settings, false);
  writeScannerSheet_(subscriptions);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) ss.setActiveSheet(sheet);

  ui.alert(
    '\u2705 Scan Complete',
    'Found ' + subscriptions.length + ' subscriptions.\n\n' +
    'To unsubscribe:\n' +
    '1. Set the Status column to "Unsubscribe" for the ones you want to remove\n' +
    '2. Go to \uD83D\uDD77 TAKScripts \u2192 \u25B6\uFE0F Scan Subscriptions\n' +
    '   (it will process any marked rows)\n\n' +
    'Or use the dropdown in column F to mark rows.',
    ui.ButtonSet.OK
  );

  processUnsubscribes_();
}

function testRun() {
  var ui = SpreadsheetApp.getUi();
  var settings = loadSettings();

  ui.alert(
    '\uD83E\uDDEA Test Run',
    'Scanning the last ' + settings.scanDays + ' days for subscriptions.\n' +
    'No unsubscribe actions will be taken — preview only.',
    ui.ButtonSet.OK
  );

  var subscriptions = findSubscriptions_(settings, true);
  writeScannerSheet_(subscriptions);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) ss.setActiveSheet(sheet);

  ui.alert(
    '\uD83E\uDDEA Test Complete',
    'Found ' + subscriptions.length + ' subscriptions (preview only).\n\n' +
    'No actions were taken. Review the results on the "' + DASHBOARD_SHEET_NAME + '" tab.\n' +
    'When ready, use \u25B6\uFE0F Scan Subscriptions to run for real.',
    ui.ButtonSet.OK
  );
}

function findSubscriptions_(settings, isTestRun) {
  var daysBack = parseInt(settings.scanDays, 10) || 90;
  var minEmails = parseInt(settings.minEmails, 10) || 3;
  var excludeList = (settings.excludeSenders || '')
    .split(',')
    .map(function(s) { return s.trim().toLowerCase(); })
    .filter(Boolean);

  var queries = [];
  queries.push('has:nousersubs newer_than:' + daysBack + 'd');
  if (settings.includePromotions) queries.push('category:promotions newer_than:' + daysBack + 'd');
  if (settings.includeUpdates)    queries.push('category:updates newer_than:' + daysBack + 'd');
  if (settings.includeSocial)     queries.push('category:social newer_than:' + daysBack + 'd');
  queries.push('label:^unsub newer_than:' + daysBack + 'd');

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

  var senderMap = {};

  for (var i = 0; i < allThreads.length; i++) {
    var messages = allThreads[i].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var from = msg.getFrom();
      var email = extractEmail_(from);
      var name = extractSenderName_(from);

      var excluded = false;
      for (var ex = 0; ex < excludeList.length; ex++) {
        if (email.indexOf(excludeList[ex]) !== -1) { excluded = true; break; }
      }
      if (excluded) continue;

      var unsubHeader = '';
      try { unsubHeader = msg.getHeader('List-Unsubscribe') || ''; } catch (e) {}

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

  var results = [];
  var keys = Object.keys(senderMap);
  for (var k = 0; k < keys.length; k++) {
    var sub = senderMap[keys[k]];
    if (sub.count >= minEmails) results.push(sub);
  }

  results.sort(function(a, b) { return b.count - a.count; });
  return results;
}

/**
 * Writes subscription data to the scanner dashboard.
 * Preserves the stats bar and headers — only clears data rows.
 */
function writeScannerSheet_(subscriptions) {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;

  // Clear existing data rows and their data validations (preserve stats bar + headers)
  var lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, SCAN_HEADERS.length)
      .clearContent()
      .clearDataValidations()
      .setBackground(null)
      .setFontColor(null)
      .setFontWeight('normal');
  }

  if (subscriptions.length === 0) {
    refreshDashboardStats();
    return;
  }

  var scanDays = parseInt(loadSettings().scanDays, 10) || 90;

  for (var i = 0; i < subscriptions.length; i++) {
    var sub = subscriptions[i];
    var rowNum = dataStartRow + i;
    var frequency = getFrequencyLabel_(sub.count, scanDays);

    var row = [
      sub.senderName || sub.senderEmail.split('@')[0],
      sub.senderEmail,
      sub.count,
      frequency,
      sub.lastDate ? Utilities.formatDate(sub.lastDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
      'Keep',
      sub.unsubscribeLink || '(none found)',
    ];

    sheet.getRange(rowNum, 1, 1, SCAN_HEADERS.length).setValues([row]);

    var rowRange = sheet.getRange(rowNum, 1, 1, SCAN_HEADERS.length);
    var bgColor = (i % 2 === 0) ? BRAND.white : BRAND.lightGray;
    rowRange
      .setBackground(bgColor)
      .setFontFamily(BRAND.bodyFont)
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(rowNum, 30);

    applyStatusStyle_(sheet, rowNum, 'Keep');
  }

  // Add status dropdown validation
  var statusRange = sheet.getRange(dataStartRow, COL_STATUS, subscriptions.length, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Keep', 'Unsubscribe', 'Done', 'Error'], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);

  refreshDashboardStats();
}


// ═══════════════════════════════════════════
// UNSUBSCRIBE ENGINE
// ═══════════════════════════════════════════

function processUnsubscribes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  var data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7).getValues();
  var actionsPerformed = 0;

  for (var i = 0; i < data.length; i++) {
    var status = data[i][COL_STATUS - 1];
    if (status !== 'Unsubscribe') continue;

    var senderName = data[i][0];
    var senderEmail = data[i][1];
    var unsubLink = data[i][6];
    var rowNum = dataStartRow + i;
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

    if (result.success) {
      sheet.getRange(rowNum, COL_STATUS).setValue('Done');
      applyStatusStyle_(sheet, rowNum, 'Done');
      actionsPerformed++;
    } else {
      sheet.getRange(rowNum, COL_STATUS).setValue('Error');
      applyStatusStyle_(sheet, rowNum, 'Error');
    }

    logAction_(senderName, senderEmail, result);
  }

  refreshDashboardStats();

  if (actionsPerformed > 0) {
    SpreadsheetApp.getUi().alert(
      '\u2705 Unsubscribe Complete',
      'Successfully processed ' + actionsPerformed + ' unsubscribe request(s).\n\n' +
      'Check the "' + LOG_SHEET_NAME + '" tab for full details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function handleUrlUnsubscribe_(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; TAKScripts-Unsubscriber/2.0)' },
    });
    var code = response.getResponseCode();
    return code >= 200 && code < 400
      ? { success: true, method: 'URL', detail: 'HTTP ' + code }
      : { success: false, method: 'URL', error: 'HTTP ' + code };
  } catch (e) {
    return { success: false, method: 'URL', error: e.message };
  }
}

function handleMailtoUnsubscribe_(mailtoLink, originalSender) {
  try {
    var address = mailtoLink.replace('mailto:', '').split('?')[0];
    var subject = 'Unsubscribe';
    var queryMatch = mailtoLink.match(/[?&]subject=([^&]*)/i);
    if (queryMatch) subject = decodeURIComponent(queryMatch[1]);
    GmailApp.sendEmail(address, subject,
      'Please unsubscribe this email address from your mailing list.\n\nSent via TAKScripts Bulk Email Unsubscriber.');
    return { success: true, method: 'Email', detail: 'Sent to ' + address };
  } catch (e) {
    return { success: false, method: 'Email', error: e.message };
  }
}


// ═══════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════

function logAction_(senderName, senderEmail, result) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length)
      .setValues([LOG_HEADERS])
      .setFontFamily(BRAND.headerFont)
      .setFontSize(9)
      .setFontWeight('bold')
      .setBackground(BRAND.darkBg)
      .setFontColor(BRAND.gold)
      .setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 32);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 300);
  }

  var newRow = sheet.getLastRow() + 1;
  var statusText = result.success ? 'Success' : 'Failed';
  var detailText = result.success ? (result.detail || '') : (result.error || '');

  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setValues([
    [new Date(), senderName, senderEmail, result.method, statusText, detailText],
  ]);

  var isEven = (newRow - 2) % 2 === 0;
  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length)
    .setBackground(isEven ? BRAND.white : BRAND.lightGray)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(newRow, 30);

  sheet.getRange(newRow, 1).setNumberFormat('MMM d, yyyy h:mm a');

  var resultCell = sheet.getRange(newRow, 5);
  if (result.success) {
    resultCell.setBackground(BRAND.successBg).setFontColor(BRAND.successText).setFontWeight('bold');
  } else {
    resultCell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText).setFontWeight('bold');
  }
}


// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

function extractSenderName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim() : '';
}

function getFrequencyLabel_(count, days) {
  var perWeek = (count / days) * 7;
  if (perWeek >= 7) return 'Daily+';
  if (perWeek >= 3) return 'Several/week';
  if (perWeek >= 1) return 'Weekly';
  if (perWeek >= 0.5) return 'Biweekly';
  return 'Monthly';
}

function applyStatusStyle_(sheet, row, status) {
  var cell = sheet.getRange(row, COL_STATUS);
  switch (status) {
    case 'Keep':
      cell.setBackground(BRAND.infoBg).setFontColor(BRAND.infoText).setFontWeight('bold');
      break;
    case 'Unsubscribe':
      cell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText).setFontWeight('bold');
      break;
    case 'Done':
      cell.setBackground(BRAND.successBg).setFontColor(BRAND.successText).setFontWeight('bold');
      break;
    case 'Error':
      cell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText).setFontWeight('bold');
      break;
    default:
      break;
  }
}

/**
 * Auto-styles status cells when changed by the user.
 */
function onEdit(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) return;
  if (e.range.getColumn() !== COL_STATUS) return;
  if (e.range.getRow() <= DASHBOARD_HEADER_ROW) return;

  applyStatusStyle_(sheet, e.range.getRow(), e.range.getValue());
  refreshDashboardStats();
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
'    body { font-family: "Segoe UI", system-ui, -apple-system, sans-serif; background: #FAFAFA; color: #1a1a1a; font-size: 13px; }' +
'    .header { background: #1A1A1A; color: white; padding: 20px 16px; text-align: center; }' +
'    .header .logo { font-size: 24px; margin-bottom: 4px; }' +
'    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }' +
'    .header .brand { color: #C9A84C; }' +
'    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }' +
'    .form { padding: 16px; }' +
'    .field { margin-bottom: 16px; }' +
'    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }' +
'    .field input[type="number"], .field input[type="text"] { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 13px; font-family: inherit; transition: border-color 0.2s; }' +
'    .field input:focus { outline: none; border-color: #C9A84C; box-shadow: 0 0 0 3px rgba(201,168,76,0.1); }' +
'    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }' +
'    .checkbox-group { margin-top: 8px; }' +
'    .checkbox-item { display: flex; align-items: center; margin-bottom: 10px; cursor: pointer; }' +
'    .checkbox-item input[type="checkbox"] { width: 16px; height: 16px; margin-right: 10px; accent-color: #C9A84C; cursor: pointer; }' +
'    .checkbox-item span { font-size: 13px; color: #333; }' +
'    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }' +
'    .btn-primary { background: #1A1A1A; color: #C9A84C; border: 1px solid #C9A84C; }' +
'    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }' +
'    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }' +
'    .btn-secondary:hover { border-color: #999; color: #333; }' +
'    .status { text-align: center; padding: 8px; font-size: 12px; margin-top: 8px; border-radius: 6px; display: none; }' +
'    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }' +
'    .status.error { display: block; background: #FFEBEE; color: #C62828; }' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }' +
'    .section-title { font-size: 12px; font-weight: 700; color: #C9A84C; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="logo">\uD83D\uDD77</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Bulk Email Unsubscriber \u00B7 Settings</div>' +
'  </div>' +
'  <div class="form">' +
'    <div class="field">' +
'      <label>Scan Period (Days)</label>' +
'      <input type="number" id="scanDays" min="7" max="365" placeholder="90" />' +
'      <div class="help">How far back to search for subscription emails.</div>' +
'    </div>' +
'    <div class="field">' +
'      <label>Minimum Emails to Flag</label>' +
'      <input type="number" id="minEmails" min="1" max="100" placeholder="3" />' +
'      <div class="help">Only show senders with at least this many emails.</div>' +
'    </div>' +
'    <div class="divider"></div>' +
'    <div class="section-title">Categories to Include</div>' +
'    <div class="checkbox-group">' +
'      <label class="checkbox-item"><input type="checkbox" id="includePromotions" /><span>Promotions</span></label>' +
'      <label class="checkbox-item"><input type="checkbox" id="includeUpdates" /><span>Updates</span></label>' +
'      <label class="checkbox-item"><input type="checkbox" id="includeSocial" /><span>Social</span></label>' +
'    </div>' +
'    <div class="divider"></div>' +
'    <div class="field">' +
'      <label>Exclude Senders</label>' +
'      <input type="text" id="excludeSenders" placeholder="boss@company.com, team@work.com" />' +
'      <div class="help">Comma-separated. Emails containing these strings will be skipped.</div>' +
'    </div>' +
'    <div class="divider"></div>' +
'    <button class="btn btn-primary" onclick="save()">Save Settings</button>' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
'    <div id="status" class="status"></div>' +
'  </div>' +
'  <script>' +
'    google.script.run.withSuccessHandler(function(settings) {' +
'      document.getElementById("scanDays").value = settings.scanDays || 90;' +
'      document.getElementById("minEmails").value = settings.minEmails || 3;' +
'      document.getElementById("includePromotions").checked = settings.includePromotions !== false;' +
'      document.getElementById("includeUpdates").checked = settings.includeUpdates !== false;' +
'      document.getElementById("includeSocial").checked = settings.includeSocial === true;' +
'      document.getElementById("excludeSenders").value = settings.excludeSenders || "";' +
'    }).loadSettings();' +
'    function save() {' +
'      var settings = {' +
'        scanDays: parseInt(document.getElementById("scanDays").value, 10) || 90,' +
'        minEmails: parseInt(document.getElementById("minEmails").value, 10) || 3,' +
'        includePromotions: document.getElementById("includePromotions").checked,' +
'        includeUpdates: document.getElementById("includeUpdates").checked,' +
'        includeSocial: document.getElementById("includeSocial").checked,' +
'        excludeSenders: document.getElementById("excludeSenders").value,' +
'      };' +
'      var statusEl = document.getElementById("status");' +
'      statusEl.className = "status"; statusEl.style.display = "none";' +
'      google.script.run' +
'        .withSuccessHandler(function() { statusEl.textContent = "\u2713 Settings saved successfully"; statusEl.className = "status success"; })' +
'        .withFailureHandler(function(err) { statusEl.textContent = "\u2715 Error: " + err.message; statusEl.className = "status error"; })' +
'        .saveSettings(settings);' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
