/**
 * Inventory Low-Stock Alert by TAKScripts
 * ========================================
 * Monitors your product inventory and sends branded alerts when
 * items drop below their reorder level.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Dashboard stats: Total · In Stock · Low Stock · Out of Stock · Alerts Today
 * - Auto-updates Status column with color-coded badges
 * - Sends branded HTML alert emails with a stock status table
 * - Optional auto-email to suppliers with reorder requests
 * - Alert history log with timestamps and actions taken
 * - onEdit trigger: instant check when Current Stock changes
 * - Scheduled checks: hourly or daily via automatic trigger
 * - Test Run: preview current status without sending alerts
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Inventory Dashboard';
var HISTORY_SHEET_NAME = '\uD83D\uDCCB Alert History';
var DASHBOARD_HEADER_ROW = 3;

var BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E8F5E9', successText: '#2E7D32',
  warningBg: '#FFF8E1', warningText: '#F57F17',
  errorBg: '#FFEBEE', errorText: '#C62828',
  headerFont: 'Roboto Mono', bodyFont: 'Roboto',
};

var HEADERS = [
  'Product Name', 'SKU', 'Category', 'Current Stock',
  'Reorder Level', 'Supplier', 'Supplier Email', 'Status',
];

var COL = {
  NAME: 1, SKU: 2, CATEGORY: 3, STOCK: 4,
  REORDER: 5, SUPPLIER: 6, SUPPLIER_EMAIL: 7, STATUS: 8,
};

var HISTORY_HEADERS = ['Timestamp', 'Product Name', 'SKU', 'Stock Level', 'Reorder Level', 'Action Taken'];

var STATUS = {
  IN_STOCK:     { label: 'In Stock',     bg: '#E8F5E9', text: '#2E7D32' },
  LOW_STOCK:    { label: 'Low Stock',    bg: '#FFF8E1', text: '#F57F17' },
  OUT_OF_STOCK: { label: 'Out of Stock', bg: '#FFEBEE', text: '#C62828' },
};


// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('\uD83D\uDD77 TAKScripts')
    .addItem('\u2699\uFE0F Settings', 'showSettings')
    .addItem('\u25B6\uFE0F Check Stock Now', 'checkStockNow')
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Dashboard', 'viewDashboard')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addItem('\uD83D\uDCCB View Alert History', 'viewAlertHistory')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Inventory Alert Settings')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function viewDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME) || getOrCreateDashboard_();
  ss.setActiveSheet(sheet);
}

function viewAlertHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('No alert history yet. Run a stock check and alerts will be logged here.');
    } catch(e) {
      Logger.log('No alert history yet. Run a stock check and alerts will be logged here.');
    }
    return;
  }
  ss.setActiveSheet(sheet);
}

function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Inventory Low-Stock Alert</h2>' +
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
  try {
    SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
  } catch(e) {
    Logger.log('About TAKScripts');
  }
}


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('inv_settings', JSON.stringify(settings));

  try {
    removeTriggers_('scheduledStockCheck');
    if (settings.frequency === 'hourly') {
      ScriptApp.newTrigger('scheduledStockCheck').timeBased().everyHours(1).create();
    } else if (settings.frequency === 'daily') {
      ScriptApp.newTrigger('scheduledStockCheck').timeBased().everyDays(1).atHour(8).create();
    }
  } catch (e) {
    Logger.log('Trigger update skipped: ' + e.message);
  }

  return { success: true };
}

function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('inv_settings');
  if (!raw) return getDefaultSettings_();
  var saved = JSON.parse(raw);
  var defaults = getDefaultSettings_();
  for (var key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}

function getDefaultSettings_() {
  return {
    alertEmail: Session.getActiveUser().getEmail() || '',
    frequency: 'onedit',
    enableSupplierEmails: false,
    thresholdMultiplier: '1.0',
  };
}


// ═══════════════════════════════════════════
// DASHBOARD MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Gets or creates the Inventory Dashboard sheet with stats bar and headers.
 * Idempotent — safe to call multiple times.
 * @return {Sheet} The dashboard sheet.
 */
function getOrCreateDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    // Migration: rename old inventory sheet if it exists
    var oldSheet = ss.getSheetByName('\uD83D\uDCE6 Inventory');
    if (oldSheet) {
      oldSheet.setName(DASHBOARD_SHEET_NAME);
      oldSheet.clear();
      sheet = oldSheet;
    } else {
      sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
    }
  } else {
    var headerCheck = sheet.getRange(DASHBOARD_HEADER_ROW, 1).getValue();
    if (headerCheck === HEADERS[0]) {
      return sheet; // Already set up
    }
  }

  var numCols = HEADERS.length;

  // ── Row 1: Stat values ──────────────────────────────────
  sheet.getRange(1, 1, 1, numCols)
    .setValues([['—', '—', '—', '—', '—', '', '', '']])
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
    .setValues([['TOTAL', 'IN STOCK', 'LOW STOCK', 'OUT OF STOCK', 'ALERTS TODAY', '', '', '']])
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
    .setValues([HEADERS])
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
  sheet.setColumnWidth(COL.NAME, 200);
  sheet.setColumnWidth(COL.SKU, 120);
  sheet.setColumnWidth(COL.CATEGORY, 140);
  sheet.setColumnWidth(COL.STOCK, 120);
  sheet.setColumnWidth(COL.REORDER, 120);
  sheet.setColumnWidth(COL.SUPPLIER, 180);
  sheet.setColumnWidth(COL.SUPPLIER_EMAIL, 220);
  sheet.setColumnWidth(COL.STATUS, 120);

  // ── Sample data ─────────────────────────────────────────
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var sample = [
    ['Widget A',       'SKU-001', 'Widgets', 50, 20, 'Acme Supply',  'orders@acme.example',  ''],
    ['Gadget B',       'SKU-002', 'Gadgets',  8, 15, 'Beta Parts',   'sales@beta.example',   ''],
    ['Gizmo C',        'SKU-003', 'Gizmos',   0, 10, 'Gamma Dist.',  'reorder@gamma.example',''],
    ['Doohickey D',    'SKU-004', 'Widgets', 100,25, 'Acme Supply',  'orders@acme.example',  ''],
    ['Thingamajig E',  'SKU-005', 'Gadgets',  3,  5, 'Beta Parts',   'sales@beta.example',   ''],
  ];
  sheet.getRange(dataStartRow, 1, sample.length, numCols)
    .setValues(sample)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10);

  return sheet;
}

/**
 * Recalculates all stats from current data and history, then updates the stats bar.
 */
function refreshDashboardStats() {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();

  var total = 0;
  var inStock = 0;
  var lowStock = 0;
  var outOfStock = 0;

  if (lastRow >= dataStartRow) {
    var numRows = lastRow - dataStartRow + 1;
    var data = sheet.getRange(dataStartRow, 1, numRows, HEADERS.length).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.NAME - 1]) continue;
      total++;
      var statusVal = row[COL.STATUS - 1];
      if (statusVal === STATUS.IN_STOCK.label)     inStock++;
      else if (statusVal === STATUS.LOW_STOCK.label)    lowStock++;
      else if (statusVal === STATUS.OUT_OF_STOCK.label) outOfStock++;
    }
  }

  // Count today's alerts from history sheet
  var alertsToday = 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (histSheet && histSheet.getLastRow() > 1) {
    var today = new Date();
    var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, 1).getValues();
    for (var h = 0; h < histData.length; h++) {
      var ts = histData[h][0];
      if (ts instanceof Date &&
          ts.getDate() === today.getDate() &&
          ts.getMonth() === today.getMonth() &&
          ts.getFullYear() === today.getFullYear()) {
        alertsToday++;
      }
    }
  }

  sheet.getRange(1, 1, 1, 5).setValues([[
    total || '—',
    inStock || '—',
    lowStock || '—',
    outOfStock || '—',
    alertsToday || '—',
  ]]);
}


// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main function — evaluates all inventory rows, updates statuses,
 * sends alerts for low/out-of-stock items.
 * @param {boolean} dryRun - If true, skips sending emails.
 * @return {Object} Summary of results for test run display.
 */
function evaluateInventory_(dryRun) {
  var sheet = getOrCreateDashboard_();
  var settings = loadSettings();
  var multiplier = parseFloat(settings.thresholdMultiplier) || 1.0;
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;

  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) {
    refreshDashboardStats();
    return { total: 0, inStock: 0, lowStock: 0, outOfStock: 0, alerts: [] };
  }

  var numRows = lastRow - dataStartRow + 1;
  var data = sheet.getRange(dataStartRow, 1, numRows, HEADERS.length).getValues();

  var counts = { total: 0, inStock: 0, lowStock: 0, outOfStock: 0 };
  var alertItems = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = row[COL.NAME - 1];
    if (!name) continue;

    var stock = Number(row[COL.STOCK - 1]) || 0;
    var reorder = Number(row[COL.REORDER - 1]) || 0;
    var threshold = Math.ceil(reorder * multiplier);
    var sku = row[COL.SKU - 1] || '';
    var supplier = row[COL.SUPPLIER - 1] || '';
    var supplierEmail = row[COL.SUPPLIER_EMAIL - 1] || '';

    var status;
    if (stock <= 0) {
      status = STATUS.OUT_OF_STOCK;
      counts.outOfStock++;
    } else if (stock <= threshold) {
      status = STATUS.LOW_STOCK;
      counts.lowStock++;
    } else {
      status = STATUS.IN_STOCK;
      counts.inStock++;
    }
    counts.total++;

    var actualRow = dataStartRow + i;

    // Alternating row background (skip status column)
    var rowBg = (i % 2 === 0) ? BRAND.white : BRAND.lightGray;
    sheet.getRange(actualRow, 1, 1, COL.STATUS - 1).setBackground(rowBg);
    sheet.setRowHeight(actualRow, 30);

    // Status cell
    sheet.getRange(actualRow, COL.STATUS)
      .setValue(status.label)
      .setBackground(status.bg)
      .setFontColor(status.text)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Highlight entire row red/amber for out of stock / low stock
    if (stock <= 0) {
      sheet.getRange(actualRow, 1, 1, COL.STATUS - 1)
        .setBackground(BRAND.errorBg);
    } else if (stock <= threshold) {
      sheet.getRange(actualRow, 1, 1, COL.STATUS - 1)
        .setBackground(BRAND.warningBg);
    }

    if (status.label !== STATUS.IN_STOCK.label) {
      alertItems.push({
        name: name,
        sku: sku,
        stock: stock,
        reorder: reorder,
        threshold: threshold,
        status: status.label,
        supplier: supplier,
        supplierEmail: supplierEmail,
      });
    }
  }

  if (!dryRun && alertItems.length > 0) {
    sendAlertEmail_(settings, alertItems);
    if (settings.enableSupplierEmails) {
      sendSupplierEmails_(alertItems);
    }
    logAlerts_(alertItems);
  }

  refreshDashboardStats();

  return {
    total: counts.total,
    inStock: counts.inStock,
    lowStock: counts.lowStock,
    outOfStock: counts.outOfStock,
    alerts: alertItems,
  };
}


// ═══════════════════════════════════════════
// EMAIL ALERTS
// ═══════════════════════════════════════════

function sendAlertEmail_(settings, alertItems) {
  var email = settings.alertEmail;
  if (!email) return;

  var subject = '\uD83D\uDD77 Inventory Alert: ' + alertItems.length + ' item(s) need attention';

  var tableRows = '';
  for (var i = 0; i < alertItems.length; i++) {
    var item = alertItems[i];
    var statusColor = item.status === 'Out of Stock' ? '#C62828' : '#F57F17';
    var statusBg = item.status === 'Out of Stock' ? '#FFEBEE' : '#FFF8E1';
    var rowBg = (i % 2 === 0) ? '#FFFFFF' : '#F9F9F9';

    tableRows +=
      '<tr style="background: ' + rowBg + ';">' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee;">' + item.name + '</td>' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #666;">' + item.sku + '</td>' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee; text-align: center; font-weight: 600;">' + item.stock + '</td>' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee; text-align: center;">' + item.reorder + '</td>' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee; text-align: center;">' +
          '<span style="background: ' + statusBg + '; color: ' + statusColor + '; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600;">' +
            item.status +
          '</span>' +
        '</td>' +
        '<td style="padding: 10px 14px; border-bottom: 1px solid #eee;">' + item.supplier + '</td>' +
      '</tr>';
  }

  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM d, yyyy \'at\' h:mm a');

  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
    '<body style="margin: 0; padding: 0; background: #f4f4f4; font-family: \'Segoe UI\', system-ui, sans-serif;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background: #f4f4f4; padding: 24px;">' +
        '<tr><td align="center">' +
          '<table width="600" cellpadding="0" cellspacing="0" style="background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">' +
            '<tr><td style="background: #1A1A1A; padding: 24px; text-align: center;">' +
              '<div style="font-size: 28px; margin-bottom: 4px;">\uD83D\uDD77</div>' +
              '<div style="color: #C9A84C; font-family: monospace; font-size: 16px; font-weight: bold; letter-spacing: 1px;">TAKScripts</div>' +
              '<div style="color: #888; font-size: 11px; margin-top: 4px;">Inventory Low-Stock Alert</div>' +
            '</td></tr>' +
            '<tr><td style="background: #FFF8E1; padding: 16px 24px; border-bottom: 2px solid #F57F17;">' +
              '<div style="font-size: 14px; font-weight: 600; color: #F57F17;">\u26A0\uFE0F ' + alertItems.length + ' item(s) require attention</div>' +
              '<div style="font-size: 12px; color: #999; margin-top: 4px;">Checked on ' + dateStr + '</div>' +
            '</td></tr>' +
            '<tr><td style="padding: 0;">' +
              '<table width="100%" cellpadding="0" cellspacing="0" style="font-size: 13px;">' +
                '<tr style="background: #1A1A1A;">' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase;">Product</td>' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase;">SKU</td>' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; text-align: center;">Stock</td>' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; text-align: center;">Reorder At</td>' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; text-align: center;">Status</td>' +
                  '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase;">Supplier</td>' +
                '</tr>' +
                tableRows +
              '</table>' +
            '</td></tr>' +
            '<tr><td style="padding: 20px 24px; text-align: center; border-top: 1px solid #eee;">' +
              '<div style="font-size: 11px; color: #999;">Sent by <span style="color: #C9A84C; font-weight: 600;">TAKScripts</span> Inventory Low-Stock Alert</div>' +
              '<div style="font-size: 11px; color: #bbb; margin-top: 4px;"><a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none;">takscripts.store</a></div>' +
            '</td></tr>' +
          '</table>' +
        '</td></tr>' +
      '</table>' +
    '</body></html>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: html, name: 'TAKScripts Inventory Alert' });
}

function sendSupplierEmails_(alertItems) {
  var bySupplier = {};
  for (var i = 0; i < alertItems.length; i++) {
    var item = alertItems[i];
    var email = item.supplierEmail;
    if (!email) continue;
    if (!bySupplier[email]) {
      bySupplier[email] = { supplier: item.supplier, items: [] };
    }
    bySupplier[email].items.push(item);
  }

  for (var supplierEmail in bySupplier) {
    var group = bySupplier[supplierEmail];
    var itemList = '';
    for (var j = 0; j < group.items.length; j++) {
      var it = group.items[j];
      itemList +=
        '<tr style="background: ' + (j % 2 === 0 ? '#FFFFFF' : '#F9F9F9') + ';">' +
          '<td style="padding: 8px 14px; border-bottom: 1px solid #eee;">' + it.name + '</td>' +
          '<td style="padding: 8px 14px; border-bottom: 1px solid #eee;">' + it.sku + '</td>' +
          '<td style="padding: 8px 14px; border-bottom: 1px solid #eee; text-align: center;">' + it.stock + '</td>' +
          '<td style="padding: 8px 14px; border-bottom: 1px solid #eee; text-align: center;">' + it.reorder + '</td>' +
        '</tr>';
    }

    var html =
      '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
      '<body style="margin: 0; padding: 0; background: #f4f4f4; font-family: \'Segoe UI\', system-ui, sans-serif;">' +
        '<table width="100%" cellpadding="0" cellspacing="0" style="background: #f4f4f4; padding: 24px;">' +
          '<tr><td align="center">' +
            '<table width="560" cellpadding="0" cellspacing="0" style="background: white; border-radius: 8px; overflow: hidden;">' +
              '<tr><td style="padding: 24px;">' +
                '<p style="font-size: 14px; color: #333;">Hi ' + group.supplier + ',</p>' +
                '<p style="font-size: 13px; color: #555; line-height: 1.6;">We\u2019d like to place a reorder for the following items that are running low:</p>' +
                '<table width="100%" cellpadding="0" cellspacing="0" style="font-size: 13px; margin: 16px 0; border: 1px solid #eee; border-radius: 6px; overflow: hidden;">' +
                  '<tr style="background: #1A1A1A;">' +
                    '<td style="padding: 8px 14px; color: #C9A84C; font-weight: bold; font-size: 11px;">Product</td>' +
                    '<td style="padding: 8px 14px; color: #C9A84C; font-weight: bold; font-size: 11px;">SKU</td>' +
                    '<td style="padding: 8px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-align: center;">Current Stock</td>' +
                    '<td style="padding: 8px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-align: center;">Reorder Level</td>' +
                  '</tr>' +
                  itemList +
                '</table>' +
                '<p style="font-size: 13px; color: #555; line-height: 1.6;">Please send us a quote or confirm availability at your earliest convenience.</p>' +
                '<p style="font-size: 13px; color: #555;">Thank you!</p>' +
              '</td></tr>' +
              '<tr><td style="padding: 12px 24px; text-align: center; border-top: 1px solid #eee;">' +
                '<span style="font-size: 10px; color: #bbb;">Sent via TAKScripts Inventory Alert</span>' +
              '</td></tr>' +
            '</table>' +
          '</td></tr>' +
        '</table>' +
      '</body></html>';

    MailApp.sendEmail({
      to: supplierEmail,
      subject: 'Reorder Request \u2014 ' + group.items.length + ' item(s)',
      htmlBody: html,
      name: 'Reorder Request',
    });
  }
}


// ═══════════════════════════════════════════
// ALERT HISTORY
// ═══════════════════════════════════════════

function logAlerts_(alertItems) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length)
      .setValues([HISTORY_HEADERS])
      .setFontFamily(BRAND.headerFont)
      .setFontSize(9)
      .setFontWeight('bold')
      .setBackground(BRAND.darkBg)
      .setFontColor(BRAND.gold)
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 250);
  }

  for (var i = 0; i < alertItems.length; i++) {
    var item = alertItems[i];
    var action = item.status === 'Out of Stock'
      ? 'Alert sent \u2014 OUT OF STOCK'
      : 'Alert sent \u2014 Low stock warning';
    if (item.supplierEmail) action += ' + Supplier notified';

    var newRow = sheet.getLastRow() + 1;
    sheet.appendRow([new Date(), item.name, item.sku, item.stock, item.reorder, action]);

    var isEven = (newRow - 2) % 2 === 0;
    sheet.getRange(newRow, 1, 1, HISTORY_HEADERS.length)
      .setBackground(isEven ? BRAND.white : BRAND.lightGray)
      .setFontFamily(BRAND.bodyFont)
      .setFontSize(10);
  }
}


// ═══════════════════════════════════════════
// TRIGGERS & CONTROLS
// ═══════════════════════════════════════════

/**
 * Fires instantly when a user changes a cell on the inventory sheet.
 * Only reacts to changes in the Current Stock column.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (col !== COL.STOCK || row <= DASHBOARD_HEADER_ROW) return;

  evaluateInventory_(false);
}

function scheduledStockCheck() {
  evaluateInventory_(false);
}

function checkStockNow() {
  var result = evaluateInventory_(false);

  var msg;
  if (result.alerts.length === 0) {
    msg = '\u2705 All ' + result.total + ' products are fully stocked. No alerts sent.';
  } else {
    msg = '\u26A0\uFE0F Stock check complete.\n\n' +
      'Total products: ' + result.total + '\n' +
      'In Stock: ' + result.inStock + '\n' +
      'Low Stock: ' + result.lowStock + '\n' +
      'Out of Stock: ' + result.outOfStock + '\n\n' +
      'Alert email sent for ' + result.alerts.length + ' item(s).';
  }

  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(msg);
  }
}

function testRun() {
  var result = evaluateInventory_(true);

  var lines = [];
  lines.push('\uD83E\uDDEA TEST RUN \u2014 No alerts will be sent');
  lines.push('');
  lines.push('Total products: ' + result.total);
  lines.push('\u2705 In Stock: ' + result.inStock);
  lines.push('\u26A0\uFE0F Low Stock: ' + result.lowStock);
  lines.push('\uD83D\uDED1 Out of Stock: ' + result.outOfStock);
  lines.push('');

  if (result.alerts.length > 0) {
    lines.push('Items that would trigger alerts:');
    lines.push('---');
    for (var i = 0; i < result.alerts.length; i++) {
      var item = result.alerts[i];
      lines.push(item.name + ' (' + item.sku + ')');
      lines.push('  Stock: ' + item.stock + ' / Reorder at: ' + item.reorder);
      lines.push('  Status: ' + item.status);
      if (item.supplierEmail) {
        lines.push('  Supplier: ' + item.supplier + ' <' + item.supplierEmail + '>');
      }
      lines.push('');
    }
  } else {
    lines.push('All products are above reorder level. No alerts needed.');
  }

  try {
    SpreadsheetApp.getUi().alert(lines.join('\n'));
  } catch(e) {
    Logger.log(lines.join('\n'));
  }
}

/**
 * First-time setup. Creates sheets and populates with sample data.
 */
function initialSetup() {
  getOrCreateDashboard_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(HISTORY_SHEET_NAME)) {
    logAlerts_([]); // Creates the history sheet with headers
  }

  evaluateInventory_(true);

  try {
    SpreadsheetApp.getUi().alert(
      '\u2705 Setup Complete!\n\n' +
      'Your inventory dashboard is ready with sample data.\n' +
      'Open \uD83D\uDD77 TAKScripts \u2192 Settings to configure alerts.\n\n' +
      'Tip: Edit the "Current Stock" column to see live status updates.'
    );
  } catch(e) {
    Logger.log('\u2705 Setup Complete!\n\nYour inventory dashboard is ready with sample data.\nOpen \uD83D\uDD77 TAKScripts \u2192 Settings to configure alerts.\n\nTip: Edit the "Current Stock" column to see live status updates.');
  }
}

function removeTriggers_(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
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
'    body { font-family: "Segoe UI", system-ui, -apple-system, sans-serif; background: #fafafa; color: #1a1a1a; font-size: 13px; }' +
'    .header { background: #1a1a1a; color: white; padding: 20px 16px; text-align: center; }' +
'    .header .logo { font-size: 24px; margin-bottom: 4px; }' +
'    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }' +
'    .header .brand { color: #C9A84C; }' +
'    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }' +
'    .form { padding: 16px; }' +
'    .field { margin-bottom: 16px; }' +
'    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }' +
'    .field input, .field select { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 13px; font-family: inherit; transition: border-color 0.2s; background: white; }' +
'    .field input:focus, .field select:focus { outline: none; border-color: #C9A84C; box-shadow: 0 0 0 3px rgba(201,168,76,0.1); }' +
'    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }' +
'    .toggle-row { display: flex; align-items: center; justify-content: space-between; padding: 12px; background: #f5f5f5; border-radius: 8px; margin-bottom: 16px; }' +
'    .toggle-row .label-text { font-size: 13px; font-weight: 500; }' +
'    .toggle-row .label-sub { font-size: 11px; color: #999; margin-top: 2px; }' +
'    .switch { position: relative; width: 44px; height: 24px; flex-shrink: 0; margin-left: 12px; }' +
'    .switch input { opacity: 0; width: 0; height: 0; }' +
'    .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background: #ccc; border-radius: 24px; transition: 0.3s; }' +
'    .slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background: white; border-radius: 50%; transition: 0.3s; }' +
'    .switch input:checked + .slider { background: #C9A84C; }' +
'    .switch input:checked + .slider:before { transform: translateX(20px); }' +
'    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }' +
'    .btn-primary { background: #1a1a1a; color: #C9A84C; border: 1px solid #C9A84C; }' +
'    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }' +
'    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }' +
'    .btn-secondary:hover { border-color: #999; color: #333; }' +
'    .status { text-align: center; padding: 8px; font-size: 12px; margin-top: 8px; border-radius: 6px; display: none; }' +
'    .status.success { display: block; background: #e8f5e9; color: #2e7d32; }' +
'    .status.error { display: block; background: #ffebee; color: #c62828; }' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }' +
'    .section-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #C9A84C; margin-bottom: 12px; }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="logo">\uD83D\uDD77</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Inventory Low-Stock Alert \u00B7 Settings</div>' +
'  </div>' +
'  <div class="form">' +
'    <div class="section-title">Notifications</div>' +
'    <div class="field">' +
'      <label>Alert Email</label>' +
'      <input type="email" id="alertEmail" placeholder="you@company.com" />' +
'      <div class="help">Low-stock alerts will be sent to this address.</div>' +
'    </div>' +
'    <div class="field">' +
'      <label>Check Frequency</label>' +
'      <select id="frequency">' +
'        <option value="onedit">On Edit (instant)</option>' +
'        <option value="hourly">Every Hour</option>' +
'        <option value="daily">Daily (8 AM)</option>' +
'      </select>' +
'      <div class="help">How often to check stock levels and send alerts.</div>' +
'    </div>' +
'    <div class="divider"></div>' +
'    <div class="section-title">Advanced</div>' +
'    <div class="field">' +
'      <label>Alert Threshold Multiplier</label>' +
'      <select id="thresholdMultiplier">' +
'        <option value="1.0">1.0x (at reorder level)</option>' +
'        <option value="1.25">1.25x (25% above reorder)</option>' +
'        <option value="1.5">1.5x (50% above reorder)</option>' +
'        <option value="2.0">2.0x (double reorder level)</option>' +
'      </select>' +
'      <div class="help">Get alerted earlier by raising the threshold above the reorder level.</div>' +
'    </div>' +
'    <div class="toggle-row">' +
'      <div>' +
'        <div class="label-text">Auto-Email Suppliers</div>' +
'        <div class="label-sub">Send reorder requests to suppliers</div>' +
'      </div>' +
'      <label class="switch">' +
'        <input type="checkbox" id="enableSupplierEmails" />' +
'        <span class="slider"></span>' +
'      </label>' +
'    </div>' +
'    <div class="divider"></div>' +
'    <div id="status" class="status"></div>' +
'    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
'  </div>' +
'  <script>' +
'    google.script.run.withSuccessHandler(function(settings) {' +
'      document.getElementById("alertEmail").value = settings.alertEmail || "";' +
'      document.getElementById("frequency").value = settings.frequency || "onedit";' +
'      document.getElementById("thresholdMultiplier").value = settings.thresholdMultiplier || "1.0";' +
'      document.getElementById("enableSupplierEmails").checked = !!settings.enableSupplierEmails;' +
'    }).loadSettings();' +
'    function save() {' +
'      var settings = {' +
'        alertEmail: document.getElementById("alertEmail").value,' +
'        frequency: document.getElementById("frequency").value,' +
'        thresholdMultiplier: document.getElementById("thresholdMultiplier").value,' +
'        enableSupplierEmails: document.getElementById("enableSupplierEmails").checked' +
'      };' +
'      var statusEl = document.getElementById("status");' +
'      var saveBtn = document.getElementById("saveBtn");' +
'      saveBtn.disabled = true;' +
'      saveBtn.textContent = "Saving\u2026";' +
'      google.script.run' +
'        .withSuccessHandler(function() {' +
'          statusEl.textContent = "\u2713 Settings saved successfully";' +
'          statusEl.className = "status success";' +
'          saveBtn.textContent = "\u2713 Saved!";' +
'          setTimeout(function() {' +
'            saveBtn.textContent = "Save Settings";' +
'            saveBtn.disabled = false;' +
'          }, 2500);' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          statusEl.textContent = "\u2715 Error: " + err.message;' +
'          statusEl.className = "status error";' +
'          saveBtn.textContent = "Save Settings";' +
'          saveBtn.disabled = false;' +
'        })' +
'        .saveSettings(settings);' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
