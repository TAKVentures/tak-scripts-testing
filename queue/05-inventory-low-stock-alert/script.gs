/**
 * Inventory Low-Stock Alert by TAKScripts
 * ========================================
 * Monitors your product inventory and sends branded alerts when
 * items drop below their reorder level.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Auto-updates Status column with color-coded badges
 * - Sends branded HTML alert emails with a stock status table
 * - Optional auto-email to suppliers with reorder requests
 * - Dashboard row showing total / in stock / low / out of stock
 * - Alert history log with timestamps and actions taken
 * - onEdit trigger: instant check when Current Stock changes
 * - Scheduled checks: hourly or daily via automatic trigger
 * - Test Run: preview current status without sending alerts
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var INVENTORY_SHEET = '\uD83D\uDCE6 Inventory';
var HISTORY_SHEET = '\uD83D\uDCCA Alert History';
var HEADERS = [
  'Product Name', 'SKU', 'Category', 'Current Stock',
  'Reorder Level', 'Supplier', 'Supplier Email', 'Status'
];
var HISTORY_HEADERS = ['Timestamp', 'Product Name', 'SKU', 'Stock Level', 'Reorder Level', 'Action Taken'];
var DATA_START_ROW = 3; // Row 1 = dashboard, Row 2 = headers
var COL = {
  NAME: 1, SKU: 2, CATEGORY: 3, STOCK: 4,
  REORDER: 5, SUPPLIER: 6, SUPPLIER_EMAIL: 7, STATUS: 8
};

// Status definitions
var STATUS = {
  IN_STOCK:     { label: 'In Stock',     bg: '#E8F5E9', text: '#2E7D32' },
  LOW_STOCK:    { label: 'Low Stock',    bg: '#FFF8E1', text: '#F57F17' },
  OUT_OF_STOCK: { label: 'Out of Stock', bg: '#FFEBEE', text: '#C62828' }
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
  ui.createMenu('\uD83D\uDD77 TAKScripts')
    .addItem('\u2699\uFE0F Settings', 'showSettings')
    .addItem('\u25B6\uFE0F Check Stock Now', 'checkStockNow')
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Alert History', 'viewAlertHistory')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Inventory Alert Settings')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Inventory Low-Stock Alert</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 \u00B7 by TAK Ventures</p>' +
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

/**
 * Navigates to the Alert History sheet.
 */
function viewAlertHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No alert history yet. Run a stock check and alerts will be logged here.');
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
  props.setProperty('inv_settings', JSON.stringify(settings));

  // Reconfigure triggers based on frequency setting
  removeTriggers_('scheduledStockCheck');
  if (settings.frequency === 'hourly') {
    ScriptApp.newTrigger('scheduledStockCheck').timeBased().everyHours(1).create();
  } else if (settings.frequency === 'daily') {
    ScriptApp.newTrigger('scheduledStockCheck').timeBased().everyDays(1).atHour(8).create();
  }
  // 'onedit' needs no time-based trigger — handled by onEdit

  return { success: true };
}

/**
 * Load saved settings.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('inv_settings');
  if (!raw) {
    return {
      alertEmail: Session.getActiveUser().getEmail() || '',
      frequency: 'onedit',
      enableSupplierEmails: false,
      thresholdMultiplier: '1.0'
    };
  }
  return JSON.parse(raw);
}

// ═══════════════════════════════════════════
// SHEET INITIALIZATION
// ═══════════════════════════════════════════

/**
 * Ensures the Inventory sheet exists with proper formatting.
 * Creates it with headers, dashboard row, and sample data if missing.
 * Idempotent — safe to call multiple times.
 */
function ensureInventorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INVENTORY_SHEET);

  if (sheet) return sheet;

  sheet = ss.insertSheet(INVENTORY_SHEET, 0);

  // ── Dashboard row (row 1) ──
  sheet.getRange(1, 1).setValue('\uD83D\uDCE6 INVENTORY DASHBOARD');
  sheet.getRange(1, 1, 1, 2).merge()
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C');

  var dashLabels = ['Total:', '0', 'In Stock:', '0', 'Low Stock:', '0', 'Out of Stock:', '0'];
  for (var d = 0; d < dashLabels.length; d++) {
    var col = d + 3;
    var cell = sheet.getRange(1, col);
    cell.setValue(dashLabels[d])
      .setFontFamily('Roboto Mono')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    if (d % 2 === 0) {
      // Label cells
      cell.setBackground('#1A1A1A').setFontColor('#C9A84C');
    } else {
      // Value cells
      var colors = {
        1: { bg: '#1A1A1A', text: '#FFFFFF' },
        3: { bg: '#E8F5E9', text: '#2E7D32' },
        5: { bg: '#FFF8E1', text: '#F57F17' },
        7: { bg: '#FFEBEE', text: '#C62828' }
      };
      cell.setBackground(colors[d].bg).setFontColor(colors[d].text);
    }
  }

  // ── Header row (row 2) ──
  sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(2);

  // ── Column widths ──
  sheet.setColumnWidth(COL.NAME, 200);
  sheet.setColumnWidth(COL.SKU, 120);
  sheet.setColumnWidth(COL.CATEGORY, 140);
  sheet.setColumnWidth(COL.STOCK, 120);
  sheet.setColumnWidth(COL.REORDER, 120);
  sheet.setColumnWidth(COL.SUPPLIER, 180);
  sheet.setColumnWidth(COL.SUPPLIER_EMAIL, 220);
  sheet.setColumnWidth(COL.STATUS, 120);

  // ── Sample data ──
  var sample = [
    ['Widget A', 'SKU-001', 'Widgets', 50, 20, 'Acme Supply', 'orders@acme.example', ''],
    ['Gadget B', 'SKU-002', 'Gadgets', 8, 15, 'Beta Parts', 'sales@beta.example', ''],
    ['Gizmo C', 'SKU-003', 'Gizmos', 0, 10, 'Gamma Dist.', 'reorder@gamma.example', ''],
    ['Doohickey D', 'SKU-004', 'Widgets', 100, 25, 'Acme Supply', 'orders@acme.example', ''],
    ['Thingamajig E', 'SKU-005', 'Gadgets', 3, 5, 'Beta Parts', 'sales@beta.example', '']
  ];
  sheet.getRange(DATA_START_ROW, 1, sample.length, HEADERS.length).setValues(sample)
    .setFontFamily('Roboto')
    .setFontSize(10);

  return sheet;
}

/**
 * Ensures the Alert History sheet exists with proper formatting.
 * Idempotent — safe to call multiple times.
 */
function ensureHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET);

  if (sheet) return sheet;

  sheet = ss.insertSheet(HISTORY_SHEET);
  sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS])
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 250);

  return sheet;
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main function — evaluates all inventory rows, updates statuses,
 * sends alerts for low/out-of-stock items.
 * @param {boolean} dryRun - If true, skips sending emails.
 * @returns {Object} Summary of results for test run display.
 */
function evaluateInventory_(dryRun) {
  var sheet = ensureInventorySheet_();
  var settings = loadSettings();
  var multiplier = parseFloat(settings.thresholdMultiplier) || 1.0;

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return { total: 0, inStock: 0, lowStock: 0, outOfStock: 0, alerts: [] };
  }

  var numRows = lastRow - DATA_START_ROW + 1;
  var data = sheet.getRange(DATA_START_ROW, 1, numRows, HEADERS.length).getValues();

  var counts = { total: 0, inStock: 0, lowStock: 0, outOfStock: 0 };
  var alertItems = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = row[COL.NAME - 1];
    if (!name) continue; // skip blank rows

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

    // Update status cell
    var statusCell = sheet.getRange(DATA_START_ROW + i, COL.STATUS);
    statusCell.setValue(status.label)
      .setBackground(status.bg)
      .setFontColor(status.text)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Apply alternating row colors (skip status column, handled above)
    var rowBg = (i % 2 === 0) ? '#FFFFFF' : '#F9F9F9';
    sheet.getRange(DATA_START_ROW + i, 1, 1, COL.STATUS - 1).setBackground(rowBg);

    // Collect low/out-of-stock items for alerting
    if (status.label !== 'In Stock') {
      alertItems.push({
        name: name,
        sku: sku,
        stock: stock,
        reorder: reorder,
        threshold: threshold,
        status: status.label,
        supplier: supplier,
        supplierEmail: supplierEmail
      });
    }
  }

  // Update dashboard
  updateDashboard_(sheet, counts);

  // Send alerts (unless dry run)
  if (!dryRun && alertItems.length > 0) {
    sendAlertEmail_(settings, alertItems);

    if (settings.enableSupplierEmails) {
      sendSupplierEmails_(alertItems);
    }

    logAlerts_(alertItems);
  }

  return {
    total: counts.total,
    inStock: counts.inStock,
    lowStock: counts.lowStock,
    outOfStock: counts.outOfStock,
    alerts: alertItems
  };
}

/**
 * Updates the dashboard row with current counts.
 */
function updateDashboard_(sheet, counts) {
  sheet.getRange(1, 4).setValue(counts.total);
  sheet.getRange(1, 6).setValue(counts.inStock);
  sheet.getRange(1, 8).setValue(counts.lowStock);
  sheet.getRange(1, 10).setValue(counts.outOfStock);
}

// ═══════════════════════════════════════════
// EMAIL ALERTS
// ═══════════════════════════════════════════

/**
 * Sends a branded HTML alert email to the configured alert email address.
 */
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
    '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"></head>' +
    '<body style="margin: 0; padding: 0; background: #f4f4f4; font-family: \'Segoe UI\', system-ui, sans-serif;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background: #f4f4f4; padding: 24px;">' +
        '<tr><td align="center">' +
          '<table width="600" cellpadding="0" cellspacing="0" style="background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">' +

            // Header
            '<tr>' +
              '<td style="background: #1A1A1A; padding: 24px; text-align: center;">' +
                '<div style="font-size: 28px; margin-bottom: 4px;">\uD83D\uDD77</div>' +
                '<div style="color: #C9A84C; font-family: monospace; font-size: 16px; font-weight: bold; letter-spacing: 1px;">TAKScripts</div>' +
                '<div style="color: #888; font-size: 11px; margin-top: 4px;">Inventory Low-Stock Alert</div>' +
              '</td>' +
            '</tr>' +

            // Alert banner
            '<tr>' +
              '<td style="background: #FFF8E1; padding: 16px 24px; border-bottom: 2px solid #F57F17;">' +
                '<div style="font-size: 14px; font-weight: 600; color: #F57F17;">' +
                  '\u26A0\uFE0F ' + alertItems.length + ' item(s) require attention' +
                '</div>' +
                '<div style="font-size: 12px; color: #999; margin-top: 4px;">Checked on ' + dateStr + '</div>' +
              '</td>' +
            '</tr>' +

            // Table
            '<tr>' +
              '<td style="padding: 0;">' +
                '<table width="100%" cellpadding="0" cellspacing="0" style="font-size: 13px;">' +
                  '<tr style="background: #1A1A1A;">' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;">Product</td>' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;">SKU</td>' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center;">Stock</td>' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center;">Reorder At</td>' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center;">Status</td>' +
                    '<td style="padding: 10px 14px; color: #C9A84C; font-weight: bold; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;">Supplier</td>' +
                  '</tr>' +
                  tableRows +
                '</table>' +
              '</td>' +
            '</tr>' +

            // Footer
            '<tr>' +
              '<td style="padding: 20px 24px; text-align: center; border-top: 1px solid #eee;">' +
                '<div style="font-size: 11px; color: #999;">Sent by <span style="color: #C9A84C; font-weight: 600;">TAKScripts</span> Inventory Low-Stock Alert</div>' +
                '<div style="font-size: 11px; color: #bbb; margin-top: 4px;">' +
                  '<a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none;">takscripts.store</a>' +
                '</div>' +
              '</td>' +
            '</tr>' +

          '</table>' +
        '</td></tr>' +
      '</table>' +
    '</body></html>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: html,
    name: 'TAKScripts Inventory Alert'
  });
}

/**
 * Sends branded reorder request emails to suppliers.
 * Groups items by supplier email to avoid duplicate messages.
 */
function sendSupplierEmails_(alertItems) {
  // Group items by supplier email
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

    var subject = 'Reorder Request \u2014 ' + group.items.length + ' item(s)';

    var html =
      '<!DOCTYPE html>' +
      '<html><head><meta charset="utf-8"></head>' +
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
              '<tr>' +
                '<td style="padding: 12px 24px; text-align: center; border-top: 1px solid #eee;">' +
                  '<span style="font-size: 10px; color: #bbb;">Sent via TAKScripts Inventory Alert</span>' +
                '</td>' +
              '</tr>' +
            '</table>' +
          '</td></tr>' +
        '</table>' +
      '</body></html>';

    MailApp.sendEmail({
      to: supplierEmail,
      subject: subject,
      htmlBody: html,
      name: 'Reorder Request'
    });
  }
}

// ═══════════════════════════════════════════
// ALERT HISTORY
// ═══════════════════════════════════════════

/**
 * Logs each alert item to the Alert History sheet.
 */
function logAlerts_(alertItems) {
  var sheet = ensureHistorySheet_();

  for (var i = 0; i < alertItems.length; i++) {
    var item = alertItems[i];
    var action = item.status === 'Out of Stock'
      ? 'Alert sent \u2014 OUT OF STOCK'
      : 'Alert sent \u2014 Low stock warning';
    if (item.supplierEmail) {
      action += ' + Supplier notified';
    }

    sheet.appendRow([
      new Date(),
      item.name,
      item.sku,
      item.stock,
      item.reorder,
      action
    ]);
  }

  // Style new rows
  var lastRow = sheet.getLastRow();
  var startRow = lastRow - alertItems.length + 1;
  if (startRow >= 2) {
    for (var r = startRow; r <= lastRow; r++) {
      var rowBg = ((r - 2) % 2 === 0) ? '#FFFFFF' : '#F9F9F9';
      sheet.getRange(r, 1, 1, HISTORY_HEADERS.length)
        .setBackground(rowBg)
        .setFontFamily('Roboto')
        .setFontSize(10);
    }
  }
}

// ═══════════════════════════════════════════
// TRIGGERS & CONTROLS
// ═══════════════════════════════════════════

/**
 * Handles edits to the Current Stock column.
 * Fires instantly when a user changes stock values.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== INVENTORY_SHEET) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Only react to changes in the Current Stock column, on data rows
  if (col !== COL.STOCK || row < DATA_START_ROW) return;

  var settings = loadSettings();
  if (settings.frequency !== 'onedit' && settings.frequency !== 'hourly' && settings.frequency !== 'daily') {
    // Default to always checking on edit
  }

  // Re-evaluate the entire inventory (updates dashboard, statuses, sends alerts)
  evaluateInventory_(false);
}

/**
 * Entry point for time-based triggers (hourly/daily).
 */
function scheduledStockCheck() {
  evaluateInventory_(false);
}

/**
 * Manual "Check Stock Now" from the menu.
 */
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

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Test run — shows current stock status without sending any alerts.
 */
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

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/**
 * Removes all triggers for a given function name.
 */
function removeTriggers_(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * First-time setup. Creates sheets and installs the onEdit trigger.
 * Called once — safe to re-run.
 */
function initialSetup() {
  ensureInventorySheet_();
  ensureHistorySheet_();

  // Apply status formatting to sample data
  evaluateInventory_(true);

  SpreadsheetApp.getUi().alert(
    '\u2705 Setup Complete!\n\n' +
    'Your inventory sheet is ready with sample data.\n' +
    'Open \uD83D\uDD77 TAKScripts \u2192 Settings to configure alerts.\n\n' +
    'Tip: Edit the "Current Stock" column to see live status updates.'
  );
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
'      background: #fafafa;' +
'      color: #1a1a1a;' +
'      font-size: 13px;' +
'    }' +
'    .header {' +
'      background: #1a1a1a;' +
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
'    .field input, .field select {' +
'      width: 100%;' +
'      padding: 10px 12px;' +
'      border: 1px solid #ddd;' +
'      border-radius: 6px;' +
'      font-size: 13px;' +
'      font-family: inherit;' +
'      transition: border-color 0.2s;' +
'      background: white;' +
'    }' +
'    .field input:focus, .field select:focus {' +
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
'    .toggle-row {' +
'      display: flex;' +
'      align-items: center;' +
'      justify-content: space-between;' +
'      padding: 12px;' +
'      background: #f5f5f5;' +
'      border-radius: 8px;' +
'      margin-bottom: 16px;' +
'    }' +
'    .toggle-row .label-text {' +
'      font-size: 13px;' +
'      font-weight: 500;' +
'    }' +
'    .toggle-row .label-sub {' +
'      font-size: 11px;' +
'      color: #999;' +
'      margin-top: 2px;' +
'    }' +
'    .switch {' +
'      position: relative;' +
'      width: 44px;' +
'      height: 24px;' +
'      flex-shrink: 0;' +
'      margin-left: 12px;' +
'    }' +
'    .switch input { opacity: 0; width: 0; height: 0; }' +
'    .slider {' +
'      position: absolute;' +
'      cursor: pointer;' +
'      top: 0; left: 0; right: 0; bottom: 0;' +
'      background: #ccc;' +
'      border-radius: 24px;' +
'      transition: 0.3s;' +
'    }' +
'    .slider:before {' +
'      position: absolute;' +
'      content: "";' +
'      height: 18px;' +
'      width: 18px;' +
'      left: 3px;' +
'      bottom: 3px;' +
'      background: white;' +
'      border-radius: 50%;' +
'      transition: 0.3s;' +
'    }' +
'    .switch input:checked + .slider { background: #C9A84C; }' +
'    .switch input:checked + .slider:before { transform: translateX(20px); }' +
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
'      background: #1a1a1a;' +
'      color: #C9A84C;' +
'      border: 1px solid #C9A84C;' +
'    }' +
'    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }' +
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
'    .status.success { display: block; background: #e8f5e9; color: #2e7d32; }' +
'    .status.error { display: block; background: #ffebee; color: #c62828; }' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }' +
'    .section-title {' +
'      font-size: 11px;' +
'      font-weight: 700;' +
'      text-transform: uppercase;' +
'      letter-spacing: 1px;' +
'      color: #C9A84C;' +
'      margin-bottom: 12px;' +
'    }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="logo">\uD83D\uDD77</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Inventory Low-Stock Alert \u00B7 Settings</div>' +
'  </div>' +
'' +
'  <div class="form">' +
'    <div class="section-title">Notifications</div>' +
'' +
'    <div class="field">' +
'      <label>Alert Email</label>' +
'      <input type="email" id="alertEmail" placeholder="you@company.com" />' +
'      <div class="help">Low-stock alerts will be sent to this address.</div>' +
'    </div>' +
'' +
'    <div class="field">' +
'      <label>Check Frequency</label>' +
'      <select id="frequency">' +
'        <option value="onedit">On Edit (instant)</option>' +
'        <option value="hourly">Every Hour</option>' +
'        <option value="daily">Daily (8 AM)</option>' +
'      </select>' +
'      <div class="help">How often to check stock levels and send alerts.</div>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'    <div class="section-title">Advanced</div>' +
'' +
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
'' +
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
'      document.getElementById("alertEmail").value = settings.alertEmail || "";' +
'      document.getElementById("frequency").value = settings.frequency || "onedit";' +
'      document.getElementById("thresholdMultiplier").value = settings.thresholdMultiplier || "1.0";' +
'      document.getElementById("enableSupplierEmails").checked = !!settings.enableSupplierEmails;' +
'    }).loadSettings();' +
'' +
'    function save() {' +
'      var settings = {' +
'        alertEmail: document.getElementById("alertEmail").value,' +
'        frequency: document.getElementById("frequency").value,' +
'        thresholdMultiplier: document.getElementById("thresholdMultiplier").value,' +
'        enableSupplierEmails: document.getElementById("enableSupplierEmails").checked' +
'      };' +
'' +
'      var statusEl = document.getElementById("status");' +
'      statusEl.className = "status";' +
'      statusEl.style.display = "none";' +
'' +
'      google.script.run' +
'        .withSuccessHandler(function() {' +
'          statusEl.textContent = "\u2713 Settings saved successfully";' +
'          statusEl.className = "status success";' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          statusEl.textContent = "\u2715 Error: " + err.message;' +
'          statusEl.className = "status error";' +
'        })' +
'        .saveSettings(settings);' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
