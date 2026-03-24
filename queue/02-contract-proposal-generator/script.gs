/**
 * Contract & Proposal Generator by TAKScripts
 * =============================================
 * Generate branded proposals and contracts from client data,
 * merge into Google Doc templates, convert to PDF, and email.
 * The Pipeline Dashboard tracks Win Rate, Total Value, Avg Deal Size,
 * Days Open per document, and flags stale proposals automatically.
 *
 * Features:
 * - Custom "🕷 TAKScripts" menu — no script editor needed
 * - Branded settings sidebar for business info & defaults
 * - Pipeline dashboard: Win Rate, Total Value, Avg Deal, Open, Stale
 * - Pulls client data from a "👥 Clients" sheet
 * - Creates polished PDFs saved to a Drive folder
 * - Optionally emails PDF to client with branded HTML email
 * - Days Open column — updates automatically on every refresh
 * - Stale proposal detection — flags docs sent 14+ days with no update
 * - Color-coded status badges: Signed (blue), Sent (amber), Expired (red), Draft (gray)
 * - Mark documents as Signed or Expired from the menu
 * - Test Run generates a doc without emailing
 * - Auto-incrementing document numbers (PROP-001, CONTRACT-001)
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
  staleBg: '#FFF8E1',   staleText: '#F57F17',
  headerFont: 'Roboto Mono',
  bodyFont: 'Roboto',
  footerText: 'Powered by TAKScripts \u00B7 takscripts.store',
};

// Dashboard layout: rows 1-2 = stats bar, row 3 = column headers, row 4+ = data
const DASHBOARD_HEADER_ROW = 3;

const CLIENT_SHEET_NAME = '\uD83D\uDC65 Clients';
const DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Pipeline Dashboard';
const FOLDER_NAME = '\uD83D\uDCC1 TAKScripts \u2014 Proposals';
const SETTINGS_KEY = 'cpg_settings';
const DOC_COUNTER_KEY = 'cpg_doc_counter';
const STALE_DAYS = 14; // Proposals sent this many days ago with no update are flagged stale

const CLIENT_HEADERS = [
  'Company', 'Contact Name', 'Email', 'Address',
  'Project Description', 'Scope', 'Deliverables',
  'Timeline', 'Rate/Price', 'Payment Terms',
];

// Dashboard columns — 8 total
const LOG_HEADERS = [
  'Doc Number', 'Client', 'Type', 'Date Created',
  'Amount', 'Status', 'Days Open', 'PDF Link',
];

// Column index map (1-based)
const COL = {
  docNumber:   1,
  client:      2,
  type:        3,
  dateCreated: 4,
  amount:      5,
  status:      6,
  daysOpen:    7,
  pdfLink:     8,
};

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
    .addSeparator()
    .addItem('\uD83D\uDCDD Generate Proposal', 'generateProposal')
    .addItem('\uD83D\uDCC4 Generate Contract', 'generateContract')
    .addSeparator()
    .addItem('\u2705 Mark as Signed', 'markAsSigned')
    .addItem('\u274C Mark as Expired', 'markAsExpired')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Dashboard', 'viewDashboard')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Contract & Proposal Settings')
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
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Contract & Proposal Generator</h2>' +
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

/**
 * Navigates to the Pipeline Dashboard sheet.
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
 * Save settings from the sidebar.
 * @param {Object} settings - Settings object from the sidebar form.
 * @return {Object} Success indicator.
 */
function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SETTINGS_KEY, JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings with sensible defaults.
 * @return {Object} The current settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SETTINGS_KEY);
  const defaults = {
    businessName: '',
    businessAddress: '',
    businessEmail: '',
    businessPhone: '',
    logoUrl: '',
    proposalPrefix: 'PROP-',
    contractPrefix: 'CONTRACT-',
    defaultTerms: 'Payment is due within 30 days of invoice date. A late fee of 1.5% per month will be applied to overdue balances. This agreement is governed by the laws of the state in which the service provider operates. Either party may terminate this agreement with 30 days written notice. All work product becomes the property of the client upon full payment.',
    defaultPaymentTerms: 'Net 30 \u2014 50% deposit required to begin work, balance due upon completion.',
    emailSubject: 'Your {{type}} from {{business}} \u2014 {{docNumber}}',
    emailBody: 'Hi {{contactName}},\n\nPlease find your {{type}} attached. Review at your convenience and let me know if you have any questions.\n\nBest regards,\n{{business}}',
  };
  if (!raw) return defaults;
  const saved = JSON.parse(raw);
  for (var key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}

// ═══════════════════════════════════════════
// DASHBOARD — THE HEART OF THE PRODUCT
// ═══════════════════════════════════════════

/**
 * Creates or gets the Pipeline Dashboard sheet with a stats bar and branded design.
 *
 * Layout:
 *   Row 1 — Stats values  (large gold numbers on dark background)
 *   Row 2 — Stats labels  (small uppercase on dark background)
 *   Row 3 — Column headers
 *   Row 4+ — Document data
 *
 * Stats: WIN RATE | TOTAL VALUE | AVG DEAL | OPEN | STALE
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The dashboard sheet.
 */
function getOrCreateDashboard_(ss) {
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);

  // ── Stats Row 1: Values ──
  sheet.getRange(1, 1, 1, 5).setValues([['—', '$0', '$0', '0', '0']]);

  // ── Stats Row 2: Labels ──
  sheet.getRange(2, 1, 1, 5).setValues([[
    'WIN RATE', 'TOTAL VALUE', 'AVG DEAL', 'OPEN', 'STALE',
  ]]);

  // Style stats values (row 1) — large bold gold
  sheet.getRange(1, 1, 1, 5)
    .setFontFamily(BRAND.headerFont)
    .setFontSize(22)
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

  // Extend dark background across all columns in stat rows
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
  sheet.setColumnWidth(COL.docNumber,   140);
  sheet.setColumnWidth(COL.client,      200);
  sheet.setColumnWidth(COL.type,        110);
  sheet.setColumnWidth(COL.dateCreated, 150);
  sheet.setColumnWidth(COL.amount,      120);
  sheet.setColumnWidth(COL.status,      110);
  sheet.setColumnWidth(COL.daysOpen,    100);
  sheet.setColumnWidth(COL.pdfLink,     120);

  // Move to first tab
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  return sheet;
}

/**
 * Refreshes the stats bar and re-applies all row styling.
 * Computes Days Open, detects stale proposals, and updates the 5 pipeline stats.
 * Called from the menu or automatically after every document log.
 */
function refreshDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('No dashboard yet. Generate a proposal or contract first.');
    } catch(e) {
      Logger.log('No dashboard yet. Generate a proposal or contract first.');
    }
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;

  if (lastRow < dataStartRow) {
    sheet.getRange(1, 1, 1, 5).setValues([['—', '$0', '$0', '0', '0']]);
    return;
  }

  const numDataRows = lastRow - dataStartRow + 1;
  const data = sheet.getRange(dataStartRow, 1, numDataRows, LOG_HEADERS.length).getValues();

  let totalValue = 0;
  let signedCount = 0;
  let concludedCount = 0; // Signed + Expired
  let openCount = 0;
  let staleCount = 0;
  let docCount = 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.docNumber - 1] || String(row[COL.docNumber - 1]).indexOf('Powered') !== -1) continue;

    docCount++;
    const actualRow = dataStartRow + i;
    const status = String(row[COL.status - 1] || '').trim();
    const statusLower = status.toLowerCase();
    const amount = parseAmount_(row[COL.amount - 1]);
    const dateCreatedVal = row[COL.dateCreated - 1];

    // Parse date created
    let dateCreated = null;
    if (dateCreatedVal) {
      dateCreated = new Date(dateCreatedVal);
      if (isNaN(dateCreated.getTime())) dateCreated = null;
    }

    // Days open
    let daysOpen = '—';
    if (dateCreated && statusLower !== 'signed' && statusLower !== 'expired') {
      dateCreated.setHours(0, 0, 0, 0);
      daysOpen = Math.floor((today - dateCreated) / (1000 * 60 * 60 * 24));
    }

    // Tally stats
    totalValue += amount;
    if (statusLower === 'signed') {
      signedCount++;
      concludedCount++;
    } else if (statusLower === 'expired') {
      concludedCount++;
    } else if (statusLower === 'sent' || statusLower === 'draft') {
      openCount++;
    }

    const isStale = statusLower === 'sent' && typeof daysOpen === 'number' && daysOpen >= STALE_DAYS;
    if (isStale) staleCount++;

    // Write Days Open value
    sheet.getRange(actualRow, COL.daysOpen).setValue(daysOpen);

    // Apply row styling
    applyRowStyle_(sheet, actualRow, statusLower, isStale);
  }

  // Compute aggregate stats
  const winRate = concludedCount > 0 ? Math.round((signedCount / concludedCount) * 100) + '%' : '—';
  const totalValueFmt = '$' + totalValue.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const avgDeal = docCount > 0 ? '$' + Math.round(totalValue / docCount).toLocaleString('en-US') : '$0';

  sheet.getRange(1, 1, 1, 5).setValues([[
    winRate,
    totalValueFmt,
    avgDeal,
    String(openCount),
    String(staleCount),
  ]]);
}

// ═══════════════════════════════════════════
// SHEET SETUP & STYLING
// ═══════════════════════════════════════════

/**
 * Ensures the Clients sheet exists with proper headers and formatting.
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The clients sheet.
 */
function ensureClientSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CLIENT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CLIENT_SHEET_NAME);
    sheet.appendRow(CLIENT_HEADERS);

    // Style header
    sheet.getRange(1, 1, 1, CLIENT_HEADERS.length)
      .setFontFamily(BRAND.headerFont)
      .setFontSize(10)
      .setFontWeight('bold')
      .setFontColor(BRAND.gold)
      .setBackground(BRAND.darkBg)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 32);
    sheet.setFrozenRows(1);

    // Column widths
    const widths = [160, 150, 220, 220, 260, 200, 200, 120, 120, 180];
    widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

    // Sample row
    sheet.appendRow([
      'Acme Corp', 'Jane Smith', 'jane@acme.com', '123 Business Ave, Suite 100',
      'Website redesign with modern UI/UX', 'Full design and development of a responsive website',
      'Wireframes, Design mockups, Developed website, Documentation',
      '6 weeks', '$5,000', 'Net 30 \u2014 50% deposit upfront',
    ]);
    styleDataRow_(sheet, 2, CLIENT_HEADERS.length);
  }

  return sheet;
}

/**
 * Applies alternating row colors and body font to a data row.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} row - The row number to style.
 * @param {number} cols - Number of columns.
 */
function styleDataRow_(sheet, row, cols) {
  const bg = row % 2 === 0 ? BRAND.lightGray : BRAND.white;
  sheet.getRange(row, 1, 1, cols)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setFontColor('#333333')
    .setBackground(bg)
    .setVerticalAlignment('middle');
}

/**
 * Applies full row styling based on document status and stale flag.
 * @param {Sheet} sheet - The dashboard sheet.
 * @param {number} row - Row number to style.
 * @param {string} statusLower - Lowercase status string.
 * @param {boolean} isStale - Whether this row is stale.
 */
function applyRowStyle_(sheet, row, statusLower, isStale) {
  const cols = LOG_HEADERS.length;

  if (isStale) {
    // Stale row — full amber highlight
    sheet.getRange(row, 1, 1, cols)
      .setBackground(BRAND.staleBg)
      .setFontFamily(BRAND.bodyFont)
      .setFontSize(10)
      .setFontColor(BRAND.staleText)
      .setVerticalAlignment('middle');
  } else {
    styleDataRow_(sheet, row, cols);
  }

  // Status cell badge (always applied regardless of row color)
  const statusCell = sheet.getRange(row, COL.status);
  statusCell.setHorizontalAlignment('center').setFontWeight('bold');
  switch (statusLower) {
    case 'signed':
      statusCell.setBackground(BRAND.infoBg).setFontColor(BRAND.infoText);
      break;
    case 'sent':
      statusCell.setBackground(isStale ? BRAND.staleBg : BRAND.warningBg)
                .setFontColor(isStale ? BRAND.staleText : BRAND.warningText);
      break;
    case 'expired':
      statusCell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText);
      break;
    case 'draft':
      statusCell.setBackground(BRAND.lightGray).setFontColor(BRAND.medGray);
      break;
    default:
      break;
  }

  // Days Open column — center align
  sheet.getRange(row, COL.daysOpen).setHorizontalAlignment('center');

  // PDF link — gold, center
  const pdfCell = sheet.getRange(row, COL.pdfLink);
  const pdfVal = pdfCell.getValue();
  if (pdfVal && String(pdfVal).indexOf('=HYPERLINK') === -1) {
    pdfCell.setFontColor(BRAND.gold).setHorizontalAlignment('center');
  }
}

// ═══════════════════════════════════════════
// CORE — DOCUMENT GENERATION
// ═══════════════════════════════════════════

/**
 * Entry point: Generate Proposal from the selected client row.
 */
function generateProposal() {
  generateDocument_('Proposal');
}

/**
 * Entry point: Generate Contract from the selected client row.
 */
function generateContract() {
  generateDocument_('Contract');
}

/**
 * Test Run — generates a proposal without emailing.
 */
function testRun() {
  generateDocument_('Proposal', true);
}

/**
 * Main generation engine. Reads client data, creates doc, converts to PDF,
 * optionally emails, and logs everything to the Pipeline Dashboard.
 * @param {string} docType - 'Proposal' or 'Contract'.
 * @param {boolean} isTest - If true, skip emailing.
 */
function generateDocument_(docType, isTest) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = loadSettings();

  if (!settings.businessName) {
    try {
      ui.alert(
        '\u26A0\uFE0F Setup Required',
        'Business name not set.\n\nOpen \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings to configure your business details.',
        ui.ButtonSet.OK
      );
    } catch(e) {
      Logger.log('Business name not set.\n\nOpen \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings to configure your business details.');
    }
    return;
  }

  const clientSheet = ensureClientSheet(ss);
  getOrCreateDashboard_(ss);

  // Must be on the Clients sheet with a row selected
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== CLIENT_SHEET_NAME) {
    try {
      ui.alert(
        '\u26A0\uFE0F Select a Client First',
        'Go to the "' + CLIENT_SHEET_NAME + '" sheet and click on the row of the client you want to generate a ' + docType.toLowerCase() + ' for.',
        ui.ButtonSet.OK
      );
    } catch(e) {
      Logger.log('Go to the "' + CLIENT_SHEET_NAME + '" sheet and click on the row of the client you want to generate a ' + docType.toLowerCase() + ' for.');
    }
    return;
  }

  const activeRow = ss.getActiveRange().getRow();
  if (activeRow < 2) {
    try {
      ui.alert('\u26A0\uFE0F Select a client row (not the header).', ui.ButtonSet.OK);
    } catch(e) {
      Logger.log('\u26A0\uFE0F Select a client row (not the header).');
    }
    return;
  }

  const rowData = clientSheet.getRange(activeRow, 1, 1, CLIENT_HEADERS.length).getValues()[0];

  const client = {
    company:      rowData[0] || '',
    contactName:  rowData[1] || '',
    email:        rowData[2] || '',
    address:      rowData[3] || '',
    description:  rowData[4] || '',
    scope:        rowData[5] || '',
    deliverables: rowData[6] || '',
    timeline:     rowData[7] || '',
    price:        rowData[8] || '',
    paymentTerms: rowData[9] || settings.defaultPaymentTerms,
  };

  if (!client.company) {
    try {
      ui.alert('\u26A0\uFE0F Missing company name in the selected row.', ui.ButtonSet.OK);
    } catch(e) {
      Logger.log('\u26A0\uFE0F Missing company name in the selected row.');
    }
    return;
  }

  const prefix = docType === 'Proposal' ? settings.proposalPrefix : settings.contractPrefix;
  const docNumber = getNextDocNumber_(prefix);
  const testLabel = isTest ? ' (TEST \u2014 no email will be sent)' : '';

  const confirm = ui.alert(
    'Generate ' + docType + testLabel,
    'Create ' + docType.toLowerCase() + ' ' + docNumber + ' for ' + client.company + '?\n\n' +
    'Contact: ' + client.contactName + '\n' +
    'Amount: ' + client.price + '\n' +
    (isTest ? '\nThis is a test run \u2014 no email will be sent.' : '\nPDF will be emailed to: ' + client.email),
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // Create Google Doc → PDF → Drive
  const doc = createDocument_(docType, docNumber, client, settings);
  const docId = doc.getId();
  const pdfBlob = convertToPdf_(docId);
  const pdfFile = savePdfToDrive_(pdfBlob, docNumber + ' \u2014 ' + client.company);
  const pdfUrl = pdfFile.getUrl();

  // Email unless test run
  let status = 'Draft';
  if (!isTest && client.email) {
    try {
      sendDocumentEmail_(docType, docNumber, client, settings, pdfBlob);
      status = 'Sent';
    } catch (e) {
      Logger.log('Email failed: ' + e.message);
      status = 'Draft';
    }
  }

  // Clean up temp doc
  DriveApp.getFileById(docId).setTrashed(true);

  // Log to dashboard
  logDocument_({
    docNumber: docNumber,
    client: client.company,
    type: docType,
    dateCreated: new Date(),
    amount: client.price,
    status: status,
    pdfUrl: pdfUrl,
  });

  const emoji = isTest ? '\uD83E\uDDEA' : '\u2705';
  try {
    ui.alert(
      emoji + ' ' + docType + ' Generated',
      docNumber + ' for ' + client.company + '\n\n' +
      'PDF saved to: ' + FOLDER_NAME +
      (isTest ? '\n\nTest run \u2014 no email was sent.' : '\n\nEmailed to: ' + client.email) +
      '\n\nView it in the \uD83D\uDCCA Pipeline Dashboard.',
      ui.ButtonSet.OK
    );
  } catch(e) {
    Logger.log(docNumber + ' for ' + client.company + '\n\nPDF saved to: ' + FOLDER_NAME + (isTest ? '\n\nTest run \u2014 no email was sent.' : '\n\nEmailed to: ' + client.email) + '\n\nView it in the \uD83D\uDCCA Pipeline Dashboard.');
  }
}

/**
 * Logs a document entry to the Pipeline Dashboard.
 * @param {Object} data - Document log entry data.
 */
function logDocument_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDashboard_(ss);

  const lastRow = sheet.getLastRow();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const newRow = Math.max(lastRow + 1, dataStartRow);

  const rowData = [
    data.docNumber,
    data.client,
    data.type,
    formatDate_(data.dateCreated),
    data.amount,
    data.status,
    0,   // Days Open — set by refreshDashboardStats
    '',  // PDF link
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  styleDataRow_(sheet, newRow, LOG_HEADERS.length);
  applyRowStyle_(sheet, newRow, data.status.toLowerCase(), false);

  // Clickable PDF link
  if (data.pdfUrl) {
    sheet.getRange(newRow, COL.pdfLink).setFormula('=HYPERLINK("' + data.pdfUrl + '","View PDF")');
    sheet.getRange(newRow, COL.pdfLink).setFontColor(BRAND.gold).setHorizontalAlignment('center');
  }

  // Refresh stats
  refreshDashboardStats();
}

// ═══════════════════════════════════════════
// STATUS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Marks the selected document row as Signed.
 */
function markAsSigned() {
  updateDocStatus_('Signed');
}

/**
 * Marks the selected document row as Expired.
 */
function markAsExpired() {
  updateDocStatus_('Expired');
}

/**
 * Updates the status of the selected row in the Pipeline Dashboard.
 * @param {string} newStatus - The new status to apply.
 */
function updateDocStatus_(newStatus) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('No dashboard found. Generate a document first.');
    } catch(e) {
      Logger.log('No dashboard found. Generate a document first.');
    }
    return;
  }

  // Must be on the dashboard sheet
  if (ss.getActiveSheet().getName() !== DASHBOARD_SHEET_NAME) {
    try {
      SpreadsheetApp.getUi().alert(
        '\u26A0\uFE0F Go to the Pipeline Dashboard first',
        'Navigate to the "\uD83D\uDCCA Pipeline Dashboard" sheet, select the row to update, then run this action again.',
        ui.ButtonSet.OK
      );
    } catch(e) {
      Logger.log('Navigate to the "\uD83D\uDCCA Pipeline Dashboard" sheet, select the row to update, then run this action again.');
    }
    return;
  }

  const row = ss.getActiveRange().getRow();
  if (row <= DASHBOARD_HEADER_ROW) {
    try {
      SpreadsheetApp.getUi().alert('\u26A0\uFE0F Select a document row below the headers.');
    } catch(e) {
      Logger.log('\u26A0\uFE0F Select a document row below the headers.');
    }
    return;
  }

  const docNumber = sheet.getRange(row, COL.docNumber).getValue();
  const client = sheet.getRange(row, COL.client).getValue();

  const confirm = ui.alert(
    'Update Status',
    'Mark ' + docNumber + ' (' + client + ') as "' + newStatus + '"?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  sheet.getRange(row, COL.status).setValue(newStatus);
  applyRowStyle_(sheet, row, newStatus.toLowerCase(), false);
  refreshDashboardStats();

  try {
    SpreadsheetApp.getUi().alert('\u2705 ' + docNumber + ' marked as ' + newStatus + '. Dashboard updated.');
  } catch(e) {
    Logger.log('\u2705 ' + docNumber + ' marked as ' + newStatus + '. Dashboard updated.');
  }
}

// ═══════════════════════════════════════════
// DOCUMENT CREATION
// ═══════════════════════════════════════════

/**
 * Creates a polished branded Google Doc for the proposal or contract.
 * @param {string} docType - 'Proposal' or 'Contract'.
 * @param {string} docNumber - Auto-incrementing document number.
 * @param {Object} client - Client data object.
 * @param {Object} settings - Current settings.
 * @return {Document} The created Google Doc.
 */
function createDocument_(docType, docNumber, client, settings) {
  const doc = DocumentApp.create(docNumber + ' \u2014 ' + client.company);
  const body = doc.getBody();

  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(50);
  body.setMarginRight(50);

  // ── Header Bar ──
  const headerTable = body.appendTable([[settings.businessName.toUpperCase()]]);
  headerTable.setBorderWidth(0);
  const headerCell = headerTable.getRow(0).getCell(0);
  headerCell.setBackgroundColor(BRAND.darkBg);
  headerCell.setPaddingTop(16);
  headerCell.setPaddingBottom(16);
  headerCell.setPaddingLeft(20);
  headerCell.setPaddingRight(20);
  const headerPara = headerCell.getChild(0).asParagraph();
  headerPara.setForegroundColor(BRAND.gold);
  headerPara.setFontFamily('Roboto Mono');
  headerPara.setFontSize(18);
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerPara.setBold(true);

  if (settings.businessAddress || settings.businessEmail || settings.businessPhone) {
    const subInfo = [settings.businessAddress, settings.businessEmail, settings.businessPhone]
      .filter(Boolean).join('  |  ');
    const subPara = headerCell.appendParagraph(subInfo);
    subPara.setForegroundColor('#999999');
    subPara.setFontFamily('Roboto');
    subPara.setFontSize(9);
    subPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  body.appendParagraph('');

  // ── Document Title ──
  const titlePara = body.appendParagraph(docType.toUpperCase());
  titlePara.setFontFamily('Roboto Mono');
  titlePara.setFontSize(22);
  titlePara.setBold(true);
  titlePara.setForegroundColor(BRAND.darkBg);

  const numberPara = body.appendParagraph(docNumber + '  \u00B7  ' + formatDate_(new Date()));
  numberPara.setFontFamily('Roboto');
  numberPara.setFontSize(10);
  numberPara.setForegroundColor('#888888');

  body.appendParagraph('');
  appendDivider_(body);

  // ── Prepared For ──
  const preparedPara = body.appendParagraph('PREPARED FOR');
  preparedPara.setFontFamily('Roboto Mono');
  preparedPara.setFontSize(9);
  preparedPara.setForegroundColor(BRAND.gold);
  preparedPara.setBold(true);
  preparedPara.setSpacingAfter(4);

  [client.company, client.contactName, client.address, client.email]
    .filter(Boolean)
    .forEach(function(line) {
      const p = body.appendParagraph(line);
      p.setFontFamily('Roboto');
      p.setFontSize(10);
      p.setForegroundColor('#333333');
      p.setSpacingAfter(2);
    });

  body.appendParagraph('');

  // ── Project Overview ──
  appendSectionHeader_(body, 'PROJECT OVERVIEW');
  if (client.description) {
    const descPara = body.appendParagraph(client.description);
    descPara.setFontFamily('Roboto');
    descPara.setFontSize(10);
    descPara.setLineSpacing(1.5);
    descPara.setForegroundColor('#333333');
  }
  body.appendParagraph('');

  // ── Scope of Work ──
  if (client.scope) {
    appendSectionHeader_(body, 'SCOPE OF WORK');
    const scopePara = body.appendParagraph(client.scope);
    scopePara.setFontFamily('Roboto');
    scopePara.setFontSize(10);
    scopePara.setLineSpacing(1.5);
    scopePara.setForegroundColor('#333333');
    body.appendParagraph('');
  }

  // ── Deliverables ──
  if (client.deliverables) {
    appendSectionHeader_(body, 'DELIVERABLES');
    const items = client.deliverables.split(',').map(function(d) { return d.trim(); }).filter(Boolean);
    const delivData = [['#', 'Deliverable']];
    items.forEach(function(item, idx) {
      delivData.push([(idx + 1).toString(), item]);
    });
    const delivTable = body.appendTable(delivData);
    styleDocTable_(delivTable);
    body.appendParagraph('');
  }

  // ── Timeline ──
  if (client.timeline) {
    appendSectionHeader_(body, 'TIMELINE');
    const tlPara = body.appendParagraph(client.timeline);
    tlPara.setFontFamily('Roboto');
    tlPara.setFontSize(10);
    tlPara.setForegroundColor('#333333');
    body.appendParagraph('');
  }

  // ── Pricing ──
  appendSectionHeader_(body, 'PRICING');
  const pricingData = [
    ['Description', 'Amount'],
    [client.description || 'Project total', client.price || 'TBD'],
  ];
  const priceTable = body.appendTable(pricingData);
  styleDocTable_(priceTable);
  priceTable.getRow(1).getCell(1).getChild(0).asParagraph().setBold(true);
  body.appendParagraph('');

  // ── Payment Terms ──
  appendSectionHeader_(body, 'PAYMENT TERMS');
  const ptPara = body.appendParagraph(client.paymentTerms || settings.defaultPaymentTerms);
  ptPara.setFontFamily('Roboto');
  ptPara.setFontSize(10);
  ptPara.setLineSpacing(1.5);
  ptPara.setForegroundColor('#333333');
  body.appendParagraph('');

  // ── Terms & Conditions ──
  appendSectionHeader_(body, 'TERMS & CONDITIONS');
  const terms = settings.defaultTerms || '';
  terms.split('. ').filter(Boolean).forEach(function(sentence) {
    const tPara = body.appendParagraph('\u2022 ' + sentence.trim() + (sentence.endsWith('.') ? '' : '.'));
    tPara.setFontFamily('Roboto');
    tPara.setFontSize(9);
    tPara.setLineSpacing(1.5);
    tPara.setForegroundColor('#555555');
  });

  body.appendParagraph('');
  appendDivider_(body);

  // ── Signatures ──
  appendSectionHeader_(body, 'ACCEPTANCE & SIGNATURES');
  const sigIntro = body.appendParagraph('By signing below, both parties agree to the terms outlined in this ' + docType.toLowerCase() + '.');
  sigIntro.setFontFamily('Roboto');
  sigIntro.setFontSize(10);
  sigIntro.setForegroundColor('#555555');
  body.appendParagraph('');

  const sigData = [
    ['FOR: ' + (settings.businessName || 'Service Provider').toUpperCase(), 'FOR: ' + client.company.toUpperCase()],
    ['', ''],
    ['Signature: ________________________________', 'Signature: ________________________________'],
    ['Name: ________________________________', 'Name: ' + client.contactName],
    ['Title: ________________________________', 'Title: ________________________________'],
    ['Date: ________________________________', 'Date: ________________________________'],
  ];
  const sigTable = body.appendTable(sigData);
  sigTable.setBorderWidth(0);
  for (let r = 0; r < sigTable.getNumRows(); r++) {
    for (let c = 0; c < sigTable.getRow(r).getNumCells(); c++) {
      const cell = sigTable.getRow(r).getCell(c);
      cell.setPaddingTop(4);
      cell.setPaddingBottom(4);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);
      const para = cell.getChild(0).asParagraph();
      para.setFontFamily('Roboto');
      para.setFontSize(9);
      para.setForegroundColor('#333333');
      if (r === 0) {
        para.setFontFamily('Roboto Mono');
        para.setFontSize(8);
        para.setBold(true);
        para.setForegroundColor(BRAND.gold);
      }
    }
  }

  body.appendParagraph('');

  // ── Doc Footer ──
  const footerTable = body.appendTable([[settings.businessName + '  \u00B7  takscripts.store  \u00B7  Generated ' + formatDate_(new Date())]]);
  footerTable.setBorderWidth(0);
  const footerCell = footerTable.getRow(0).getCell(0);
  footerCell.setBackgroundColor('#F5F5F5');
  footerCell.setPaddingTop(10);
  footerCell.setPaddingBottom(10);
  footerCell.setPaddingLeft(16);
  footerCell.setPaddingRight(16);
  const footerPara = footerCell.getChild(0).asParagraph();
  footerPara.setForegroundColor('#999999');
  footerPara.setFontFamily('Roboto');
  footerPara.setFontSize(8);
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();
  return doc;
}

// ═══════════════════════════════════════════
// DOCUMENT HELPERS
// ═══════════════════════════════════════════

/**
 * Appends a branded gold section header to a document body.
 * @param {Body} body - The document body.
 * @param {string} text - Header text.
 */
function appendSectionHeader_(body, text) {
  const para = body.appendParagraph(text);
  para.setFontFamily('Roboto Mono');
  para.setFontSize(10);
  para.setBold(true);
  para.setForegroundColor(BRAND.darkBg);
  para.setSpacingAfter(6);
  para.setSpacingBefore(8);
}

/**
 * Appends a thin gold divider line to a document body.
 * @param {Body} body - The document body.
 */
function appendDivider_(body) {
  const divTable = body.appendTable([['']]);
  divTable.setBorderWidth(0);
  const divCell = divTable.getRow(0).getCell(0);
  divCell.setBackgroundColor(BRAND.gold);
  divCell.setPaddingTop(0);
  divCell.setPaddingBottom(0);
  divCell.getChild(0).asParagraph().setFontSize(1).setText('');
}

/**
 * Styles a document table with a branded dark header and alternating data rows.
 * @param {Table} table - The Google Doc table to style.
 */
function styleDocTable_(table) {
  table.setBorderWidth(1);
  table.setBorderColor(BRAND.border);

  const headerRow = table.getRow(0);
  for (let c = 0; c < headerRow.getNumCells(); c++) {
    const cell = headerRow.getCell(c);
    cell.setBackgroundColor(BRAND.darkBg);
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(10);
    cell.setPaddingRight(10);
    const para = cell.getChild(0).asParagraph();
    para.setForegroundColor(BRAND.gold);
    para.setFontFamily('Roboto Mono');
    para.setFontSize(9);
    para.setBold(true);
  }

  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const bgColor = r % 2 === 0 ? BRAND.lightGray : BRAND.white;
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor(bgColor);
      cell.setPaddingTop(6);
      cell.setPaddingBottom(6);
      cell.setPaddingLeft(10);
      cell.setPaddingRight(10);
      const para = cell.getChild(0).asParagraph();
      para.setFontFamily('Roboto');
      para.setFontSize(10);
      para.setForegroundColor('#333333');
    }
  }
}

// ═══════════════════════════════════════════
// PDF & DRIVE
// ═══════════════════════════════════════════

/**
 * Converts a Google Doc to a PDF blob via export URL.
 * @param {string} docId - The Google Doc ID.
 * @return {Blob} The PDF blob.
 */
function convertToPdf_(docId) {
  const url = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  return response.getBlob().setName(DriveApp.getFileById(docId).getName() + '.pdf');
}

/**
 * Saves a PDF to the TAKScripts Proposals folder. Creates folder if needed.
 * @param {Blob} pdfBlob - The PDF blob.
 * @param {string} fileName - The file name (without extension).
 * @return {File} The saved Drive file.
 */
function savePdfToDrive_(pdfBlob, fileName) {
  let folder;
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }
  pdfBlob.setName(fileName + '.pdf');
  return folder.createFile(pdfBlob);
}

// ═══════════════════════════════════════════
// EMAIL
// ═══════════════════════════════════════════

/**
 * Sends the document as a branded HTML email with PDF attached.
 * @param {string} docType - 'Proposal' or 'Contract'.
 * @param {string} docNumber - Document number.
 * @param {Object} client - Client data object.
 * @param {Object} settings - Current settings.
 * @param {Blob} pdfBlob - The PDF to attach.
 */
function sendDocumentEmail_(docType, docNumber, client, settings, pdfBlob) {
  const subject = (settings.emailSubject || 'Your {{type}} \u2014 {{docNumber}}')
    .replace(/\{\{type\}\}/g, docType.toLowerCase())
    .replace(/\{\{business\}\}/g, settings.businessName)
    .replace(/\{\{docNumber\}\}/g, docNumber)
    .replace(/\{\{contactName\}\}/g, client.contactName || 'there');

  const bodyText = (settings.emailBody || '')
    .replace(/\{\{type\}\}/g, docType.toLowerCase())
    .replace(/\{\{business\}\}/g, settings.businessName)
    .replace(/\{\{docNumber\}\}/g, docNumber)
    .replace(/\{\{contactName\}\}/g, client.contactName || 'there');

  const htmlBody = getEmailHtml_(docType, docNumber, client, settings, bodyText);

  GmailApp.sendEmail(client.email, subject, bodyText, {
    htmlBody: htmlBody,
    attachments: [pdfBlob],
    name: settings.businessName || 'TAKScripts',
  });
}

/**
 * Returns the branded HTML email template.
 * @param {string} docType - 'Proposal' or 'Contract'.
 * @param {string} docNumber - Document number.
 * @param {Object} client - Client data object.
 * @param {Object} settings - Current settings.
 * @param {string} bodyText - Plain text body for HTML conversion.
 * @return {string} HTML string.
 */
function getEmailHtml_(docType, docNumber, client, settings, bodyText) {
  const bodyHtml = bodyText.replace(/\n/g, '<br>');
  const businessName = escapeHtml_(settings.businessName || 'TAKScripts');

  return '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:\'Segoe UI\',system-ui,sans-serif;">' +
    '<div style="max-width:600px;margin:0 auto;background:#ffffff;">' +
    // Header
    '<div style="background:#1A1A1A;padding:28px 32px;text-align:center;">' +
    '<div style="font-size:28px;margin-bottom:4px;">\uD83D\uDD77</div>' +
    '<div style="color:#C9A84C;font-size:16px;font-weight:700;letter-spacing:1px;">' + businessName.toUpperCase() + '</div>' +
    '</div>' +
    // Body
    '<div style="padding:32px;line-height:1.7;color:#333333;font-size:14px;">' +
    bodyHtml +
    '<div style="margin-top:24px;padding:16px;background:#F9F9F9;border-radius:8px;border-left:3px solid #C9A84C;">' +
    '<div style="font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Document Details</div>' +
    '<div style="font-size:13px;color:#333;">' +
    '<strong>' + escapeHtml_(docType) + ':</strong> ' + escapeHtml_(docNumber) + '<br>' +
    '<strong>Client:</strong> ' + escapeHtml_(client.company) + '<br>' +
    '<strong>Amount:</strong> ' + escapeHtml_(client.price || 'See attached') +
    '</div></div></div>' +
    // Footer
    '<div style="background:#F5F5F5;padding:20px 32px;text-align:center;border-top:1px solid #eee;">' +
    '<div style="font-size:11px;color:#999;line-height:1.6;">' +
    businessName +
    (settings.businessAddress ? '<br>' + escapeHtml_(settings.businessAddress) : '') +
    (settings.businessPhone ? '<br>' + escapeHtml_(settings.businessPhone) : '') +
    '<br><br><span style="color:#C9A84C;">' + BRAND.footerText + '</span>' +
    '</div></div></div></body></html>';
}

// ═══════════════════════════════════════════
// DOCUMENT NUMBERING
// ═══════════════════════════════════════════

/**
 * Returns the next auto-incrementing document number for the given prefix.
 * @param {string} prefix - The document prefix (e.g., 'PROP-').
 * @return {string} The formatted document number (e.g., 'PROP-001').
 */
function getNextDocNumber_(prefix) {
  const props = PropertiesService.getScriptProperties();
  const counterKey = DOC_COUNTER_KEY + '_' + prefix;
  const current = parseInt(props.getProperty(counterKey) || '0', 10);
  const next = current + 1;
  props.setProperty(counterKey, next.toString());
  return prefix + padNumber_(next, 3);
}

/**
 * Pads a number with leading zeros to a fixed length.
 * @param {number} num - The number to pad.
 * @param {number} length - The desired string length.
 * @return {string} Zero-padded number string.
 */
function padNumber_(num, length) {
  let str = num.toString();
  while (str.length < length) str = '0' + str;
  return str;
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Formats a Date object to a readable string like "March 22, 2026".
 * @param {Date} date - The date to format.
 * @return {string} Formatted date string.
 */
function formatDate_(date) {
  const months = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

/**
 * Parses an amount string like "$5,000" or "5000" into a number.
 * @param {*} val - Raw cell value.
 * @return {number} Parsed numeric amount, or 0 if unparseable.
 */
function parseAmount_(val) {
  if (!val) return 0;
  const cleaned = String(val).replace(/[$,\s]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * Escapes HTML special characters to prevent injection in email templates.
 * @param {string} text - Raw text.
 * @return {string} HTML-safe text.
 */
function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ═══════════════════════════════════════════
// FIRST-RUN SETUP
// ═══════════════════════════════════════════

/**
 * One-time setup: creates all sheets with proper formatting.
 * Safe to run multiple times — idempotent.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureClientSheet(ss);
  getOrCreateDashboard_(ss);
  ss.setActiveSheet(ss.getSheetByName(CLIENT_SHEET_NAME));

  try {
    SpreadsheetApp.getUi().alert(
      '\u2705 Setup Complete',
      'Your sheets are ready:\n\n' +
      '\u2022 \uD83D\uDCCA Pipeline Dashboard \u2014 Win rate, totals, and document log\n' +
      '\u2022 \uD83D\uDC65 Clients \u2014 Add your client data here\n\n' +
      'Next: Go to \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings to add your business details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    Logger.log('Your sheets are ready. Next: Go to \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings to add your business details.');
  }
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

/**
 * Returns the full HTML for the settings sidebar.
 * Branded with TAKScripts design system.
 * @return {string} HTML string.
 */
function getSettingsHtml() {
  return `<!DOCTYPE html>
<html>
<head>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: #fafafa;
      color: #1a1a1a;
      font-size: 13px;
    }
    .header {
      background: #1a1a1a;
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
    .section-title:first-child { margin-top: 4px; }
    .field { margin-bottom: 14px; }
    .field label {
      display: block;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #666;
      margin-bottom: 5px;
    }
    .field input, .field textarea {
      width: 100%;
      padding: 9px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      font-family: inherit;
      transition: border-color 0.2s, box-shadow 0.2s;
      background: white;
    }
    .field input:focus, .field textarea:focus {
      outline: none;
      border-color: #C9A84C;
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea { min-height: 80px; resize: vertical; line-height: 1.5; }
    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
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
    .btn-primary { background: #1a1a1a; color: #C9A84C; border: 1px solid #C9A84C; }
    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }
    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .btn-action {
      background: white;
      color: #333;
      border: 1px solid #ddd;
      margin-top: 6px;
      font-size: 12px;
      padding: 10px;
    }
    .btn-action:hover { border-color: #C9A84C; color: #C9A84C; }
    .status {
      text-align: center;
      padding: 8px;
      font-size: 12px;
      margin-top: 8px;
      border-radius: 6px;
      display: none;
    }
    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }
    .status.error { display: block; background: #FFEBEE; color: #C62828; }
    .divider { border-top: 1px solid #eee; margin: 20px 0; }
    .variables {
      background: #f5f5f5;
      border-radius: 6px;
      padding: 10px 12px;
      margin-top: 6px;
    }
    .variables code {
      display: inline-block;
      background: #e8e8e8;
      padding: 1px 6px;
      border-radius: 3px;
      font-size: 11px;
      margin: 2px;
    }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo">\uD83D\uDD77</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">Contract & Proposal Generator \u00B7 Settings</div>
  </div>

  <div class="form">

    <div class="section-title">Business Information</div>

    <div class="field">
      <label>Business Name</label>
      <input type="text" id="businessName" placeholder="Your Company Name" />
    </div>
    <div class="field">
      <label>Address</label>
      <input type="text" id="businessAddress" placeholder="123 Main St, City, State ZIP" />
    </div>
    <div class="row">
      <div class="field">
        <label>Email</label>
        <input type="email" id="businessEmail" placeholder="you@company.com" />
      </div>
      <div class="field">
        <label>Phone</label>
        <input type="tel" id="businessPhone" placeholder="(555) 123-4567" />
      </div>
    </div>
    <div class="field">
      <label>Logo URL (optional)</label>
      <input type="url" id="logoUrl" placeholder="https://..." />
      <div class="help">Direct link to your logo image. Used in documents.</div>
    </div>

    <div class="section-title">Document Numbering</div>

    <div class="row">
      <div class="field">
        <label>Proposal Prefix</label>
        <input type="text" id="proposalPrefix" placeholder="PROP-" />
      </div>
      <div class="field">
        <label>Contract Prefix</label>
        <input type="text" id="contractPrefix" placeholder="CONTRACT-" />
      </div>
    </div>

    <div class="section-title">Default Terms</div>

    <div class="field">
      <label>Terms & Conditions</label>
      <textarea id="defaultTerms" rows="4" placeholder="Payment is due within 30 days..."></textarea>
      <div class="help">Appears in every proposal and contract.</div>
    </div>
    <div class="field">
      <label>Payment Terms</label>
      <textarea id="defaultPaymentTerms" rows="2" placeholder="Net 30 — 50% deposit required..."></textarea>
    </div>

    <div class="section-title">Email Template</div>

    <div class="field">
      <label>Email Subject</label>
      <input type="text" id="emailSubject" placeholder="Your {{type}} from {{business}}" />
    </div>
    <div class="field">
      <label>Email Body</label>
      <textarea id="emailBody" rows="4" placeholder="Hi {{contactName}},"></textarea>
      <div class="variables">
        <div class="help" style="margin-bottom: 4px;">Template variables:</div>
        <code>{{contactName}}</code> <code>{{type}}</code>
        <code>{{business}}</code> <code>{{docNumber}}</code>
      </div>
    </div>

    <div id="status" class="status"></div>
    <div class="divider"></div>

    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div class="divider"></div>

    <div class="section-title">Quick Actions</div>
    <button class="btn btn-action" onclick="runAction('markAsSigned')">\u2705 Mark Selected Row as Signed</button>
    <button class="btn btn-action" onclick="runAction('markAsExpired')">\u274C Mark Selected Row as Expired</button>
    <button class="btn btn-action" onclick="runAction('refreshDashboardStats')" style="margin-top: 10px;">\uD83D\uDD04 Refresh Dashboard Stats</button>
    <button class="btn btn-action" onclick="runAction('setupSheets')">\uD83D\uDD27 Run Initial Setup</button>
  </div>

  <script>
    google.script.run.withSuccessHandler(function(s) {
      document.getElementById('businessName').value = s.businessName || '';
      document.getElementById('businessAddress').value = s.businessAddress || '';
      document.getElementById('businessEmail').value = s.businessEmail || '';
      document.getElementById('businessPhone').value = s.businessPhone || '';
      document.getElementById('logoUrl').value = s.logoUrl || '';
      document.getElementById('proposalPrefix').value = s.proposalPrefix || 'PROP-';
      document.getElementById('contractPrefix').value = s.contractPrefix || 'CONTRACT-';
      document.getElementById('defaultTerms').value = s.defaultTerms || '';
      document.getElementById('defaultPaymentTerms').value = s.defaultPaymentTerms || '';
      document.getElementById('emailSubject').value = s.emailSubject || '';
      document.getElementById('emailBody').value = s.emailBody || '';
    }).loadSettings();

    function save() {
      var settings = {
        businessName: document.getElementById('businessName').value,
        businessAddress: document.getElementById('businessAddress').value,
        businessEmail: document.getElementById('businessEmail').value,
        businessPhone: document.getElementById('businessPhone').value,
        logoUrl: document.getElementById('logoUrl').value,
        proposalPrefix: document.getElementById('proposalPrefix').value || 'PROP-',
        contractPrefix: document.getElementById('contractPrefix').value || 'CONTRACT-',
        defaultTerms: document.getElementById('defaultTerms').value,
        defaultPaymentTerms: document.getElementById('defaultPaymentTerms').value,
        emailSubject: document.getElementById('emailSubject').value,
        emailBody: document.getElementById('emailBody').value,
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

    function runAction(fnName) {
      showStatus('', '');
      google.script.run
        .withSuccessHandler(function() { showStatus('\u2713 Done', 'success'); })
        .withFailureHandler(function(err) { showStatus('\u2715 Error: ' + err.message, 'error'); })
        [fnName]();
    }

    function showStatus(msg, type) {
      var el = document.getElementById('status');
      if (!msg) { el.style.display = 'none'; return; }
      el.textContent = msg;
      el.className = 'status ' + type;
    }
  </script>
</body>
</html>`;
}
