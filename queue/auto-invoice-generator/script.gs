/**
 * Auto Invoice Generator by TAKScripts
 * ======================================
 * Generate branded PDF invoices, email them to clients, and track everything
 * from a single Google Sheet — no code required.
 *
 * Features:
 * - Custom menu in Google Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for business info, tax, and email config
 * - Reads client data from a styled "Clients" sheet
 * - Generates branded Google Doc invoices from an auto-created template
 * - Converts to PDF and saves to an organized Drive folder
 * - Sends professional HTML emails with PDF attachment
 * - Logs every invoice with status tracking (Draft / Sent / Paid)
 * - Configurable invoice prefix and auto-incrementing numbers
 * - Test Run mode — generates a sample without emailing
 * - Mark invoices as Paid directly from the log
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const BRAND = {
  headerBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  bodyBg: '#FAFAFA',
  successBg: '#E8F5E9',
  successText: '#2E7D32',
  warningBg: '#FFF8E1',
  warningText: '#F57F17',
  errorBg: '#FFEBEE',
  errorText: '#C62828',
  headerFont: 'Roboto Mono',
  bodyFont: 'Roboto',
  headerSize: 10,
  bodySize: 10,
  footerText: 'Powered by TAKScripts \u00B7 takscripts.store',
};

const SHEET_NAMES = {
  clients: '\uD83D\uDC65 Clients',
  log: '\uD83D\uDCCA Invoice Log',
};

const FOLDER_NAME = '\uD83D\uDCC1 TAKScripts \u2014 Invoices';

const CLIENT_HEADERS = [
  'Client Name', 'Email', 'Address', 'Service Description',
  'Rate ($)', 'Hours/Qty', 'Notes',
];

const LOG_HEADERS = [
  'Invoice #', 'Client Name', 'Email', 'Amount ($)',
  'Date Created', 'Due Date', 'Status', 'PDF Link',
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
    .addItem('\u25B6\uFE0F Generate Invoice', 'generateInvoicePrompt')
    .addItem('\uD83D\uDCCA View Invoice Log', 'viewInvoiceLog')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Invoice Generator Settings')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Segoe UI', system-ui, sans-serif; padding: 20px; text-align: center;">
      <div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>
      <h2 style="margin: 0 0 4px; font-size: 18px;">Auto Invoice Generator</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 \u00B7 by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
        Part of the <strong>TAKScripts</strong> collection.<br>
        Pre-built Google Apps Scripts for small business.
      </p>
      <p style="margin-top: 16px;">
        <a href="https://takscripts.store" target="_blank"
           style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">
          takscripts.store \u2192
        </a>
      </p>
    </div>
  `).setWidth(300).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to or creates the Invoice Log sheet.
 */
function viewInvoiceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.log);
  if (!sheet) {
    sheet = createLogSheet_(ss);
  }
  ss.setActiveSheet(sheet);
}

// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Default settings for a fresh install.
 */
function getDefaultSettings_() {
  return {
    businessName: '',
    businessAddress: '',
    businessEmail: '',
    logoUrl: '',
    taxRate: 0,
    paymentTerms: 'Payment due within the specified number of days. Late payments may be subject to a fee.',
    dueDays: 30,
    invoicePrefix: 'INV-',
    nextInvoiceNumber: 1001,
    emailSubject: 'Invoice {{invoice_number}} from {{business_name}}',
    emailBody: 'Hi {{client_name}},\n\nPlease find attached invoice {{invoice_number}} for {{total}}.\n\nPayment is due by {{due_date}}.\n\nThank you for your business!\n\nBest regards,\n{{business_name}}',
    autoEmail: false,
  };
}

/**
 * Save settings from the sidebar.
 * @param {Object} settings - Settings object from the sidebar form.
 * @return {Object} Success indicator.
 */
function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('invoice_settings', JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings, merging with defaults for any missing keys.
 * @return {Object} The current settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('invoice_settings');
  const defaults = getDefaultSettings_();
  if (!raw) return defaults;
  const saved = JSON.parse(raw);
  // Merge so new setting keys are always present
  return Object.assign({}, defaults, saved);
}

// ═══════════════════════════════════════════
// SHEET SETUP & STYLING
// ═══════════════════════════════════════════

/**
 * Creates and styles the Clients sheet with branded headers.
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The newly created sheet.
 */
function createClientsSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_NAMES.clients);
  sheet.appendRow(CLIENT_HEADERS);

  // Style header row
  const headerRange = sheet.getRange(1, 1, 1, CLIENT_HEADERS.length);
  headerRange
    .setFontFamily(BRAND.headerFont)
    .setFontSize(BRAND.headerSize)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.headerBg)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Column widths
  const widths = [200, 250, 280, 250, 100, 100, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Add sample row so users see the expected format
  const sampleRow = [
    'Acme Corp',
    'billing@acme.com',
    '123 Main St, Springfield, IL 62701',
    'Website Development',
    150,
    10,
    'Monthly retainer',
  ];
  sheet.appendRow(sampleRow);
  styleDataRow_(sheet, 2, CLIENT_HEADERS.length);

  // Footer
  addFooter_(sheet, CLIENT_HEADERS.length);

  return sheet;
}

/**
 * Creates and styles the Invoice Log sheet with branded headers.
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The newly created sheet.
 */
function createLogSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_NAMES.log);
  sheet.appendRow(LOG_HEADERS);

  // Style header row
  const headerRange = sheet.getRange(1, 1, 1, LOG_HEADERS.length);
  headerRange
    .setFontFamily(BRAND.headerFont)
    .setFontSize(BRAND.headerSize)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.headerBg)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Column widths
  const widths = [120, 200, 250, 120, 140, 140, 110, 280];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Footer
  addFooter_(sheet, LOG_HEADERS.length);

  return sheet;
}

/**
 * Applies alternating row colors and body font to a data row.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} row - The row number to style.
 * @param {number} cols - Number of columns.
 */
function styleDataRow_(sheet, row, cols) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setFontFamily(BRAND.bodyFont).setFontSize(BRAND.bodySize);
  const bg = row % 2 === 0 ? BRAND.lightGray : BRAND.white;
  range.setBackground(bg);
}

/**
 * Applies a status badge style to a cell based on status value.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} row - Row number.
 * @param {number} col - Column number.
 * @param {string} status - The status text (Sent, Draft, Paid, Error).
 */
function styleStatusCell_(sheet, row, col, status) {
  const cell = sheet.getRange(row, col);
  cell.setHorizontalAlignment('center').setFontWeight('bold');
  switch (status.toLowerCase()) {
    case 'paid':
      cell.setBackground(BRAND.successBg).setFontColor(BRAND.successText);
      break;
    case 'sent':
      cell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText);
      break;
    case 'draft':
      cell.setBackground(BRAND.lightGray).setFontColor('#666666');
      break;
    case 'error':
      cell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText);
      break;
    default:
      break;
  }
}

/**
 * Adds the branded footer row to a sheet.
 * @param {Sheet} sheet - The target sheet.
 * @param {number} cols - Number of columns to merge across.
 */
function addFooter_(sheet, cols) {
  // Leave a gap row, then add footer
  const lastRow = sheet.getLastRow() + 2;
  const footerRange = sheet.getRange(lastRow, 1, 1, cols);
  footerRange.merge();
  footerRange
    .setValue(BRAND.footerText)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(9)
    .setFontColor('#999999')
    .setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════
// INVOICE GENERATION — CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Prompts the user to select which client row to invoice.
 * Called from the menu.
 */
function generateInvoicePrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure Clients sheet exists
  let clientSheet = ss.getSheetByName(SHEET_NAMES.clients);
  if (!clientSheet) {
    clientSheet = createClientsSheet_(ss);
    ss.setActiveSheet(clientSheet);
    ui.alert(
      'Setup Required',
      'A "' + SHEET_NAMES.clients + '" sheet has been created.\n\n' +
      'Fill in your client data and run Generate Invoice again.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Check for settings
  const settings = loadSettings();
  if (!settings.businessName) {
    ui.alert(
      'Setup Required',
      'Please configure your business details first.\n\n' +
      'Go to: \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings',
      ui.ButtonSet.OK
    );
    return;
  }

  // Get client data
  const data = clientSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('No Clients', 'Add at least one client to the "' + SHEET_NAMES.clients + '" sheet first.', ui.ButtonSet.OK);
    return;
  }

  // Build a client list for the prompt
  const clientNames = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    if (name && String(name).trim()) {
      clientNames.push((i) + '. ' + name);
    }
  }

  if (clientNames.length === 0) {
    ui.alert('No Clients', 'Add at least one client to the "' + SHEET_NAMES.clients + '" sheet first.', ui.ButtonSet.OK);
    return;
  }

  const result = ui.prompt(
    'Generate Invoice',
    'Enter the row number of the client to invoice:\n\n' + clientNames.join('\n') +
    '\n\n(Or type "all" to invoice every client)',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const input = result.getResponseText().trim().toLowerCase();

  if (input === 'all') {
    // Invoice all clients
    let count = 0;
    let errors = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        try {
          generateInvoiceForRow_(data[i], settings, false);
          count++;
        } catch (e) {
          Logger.log('Error invoicing row ' + (i + 1) + ': ' + e.message);
          errors++;
        }
      }
    }
    ui.alert('Batch Complete', count + ' invoice(s) generated.' + (errors > 0 ? '\n' + errors + ' error(s) — check the log.' : ''), ui.ButtonSet.OK);
  } else {
    const rowIndex = parseInt(input, 10);
    if (isNaN(rowIndex) || rowIndex < 1 || rowIndex >= data.length) {
      ui.alert('Invalid row number. Please try again.');
      return;
    }
    try {
      const invoice = generateInvoiceForRow_(data[rowIndex], settings, false);
      ui.alert(
        'Invoice Generated',
        'Invoice ' + invoice.invoiceNumber + ' created for ' + invoice.clientName + '.\n' +
        'Amount: $' + invoice.total.toFixed(2) + '\n' +
        (invoice.emailed ? 'Email sent to ' + invoice.clientEmail : 'Saved as draft (email not sent)') + '\n\n' +
        'PDF saved to Drive folder: ' + FOLDER_NAME,
        ui.ButtonSet.OK
      );
    } catch (e) {
      ui.alert('Error', 'Failed to generate invoice:\n' + e.message, ui.ButtonSet.OK);
      Logger.log('Invoice generation error: ' + e.stack);
    }
  }
}

/**
 * Generates an invoice for a single client row.
 * @param {Array} rowData - Array of cell values from the Clients sheet.
 * @param {Object} settings - Current settings.
 * @param {boolean} isTest - If true, skip emailing and mark as Draft.
 * @return {Object} Invoice summary with key details.
 */
function generateInvoiceForRow_(rowData, settings, isTest) {
  const clientName = String(rowData[0]).trim();
  const clientEmail = String(rowData[1]).trim();
  const clientAddress = String(rowData[2]).trim();
  const serviceDesc = String(rowData[3]).trim();
  const rate = parseFloat(rowData[4]) || 0;
  const hours = parseFloat(rowData[5]) || 0;
  const notes = String(rowData[6] || '').trim();

  if (!clientName) throw new Error('Client name is required.');
  if (!clientEmail && !isTest) throw new Error('Client email is required for ' + clientName + '.');

  // Calculate amounts
  const subtotal = rate * hours;
  const taxRate = parseFloat(settings.taxRate) || 0;
  const taxAmount = subtotal * (taxRate / 100);
  const total = subtotal + taxAmount;

  // Generate invoice number
  const invoiceNumber = settings.invoicePrefix + settings.nextInvoiceNumber;

  // Dates
  const today = new Date();
  const dueDays = parseInt(settings.dueDays, 10) || 30;
  const dueDate = new Date(today);
  dueDate.setDate(dueDate.getDate() + dueDays);

  // Create the invoice document
  const docId = createInvoiceDoc_({
    invoiceNumber: invoiceNumber,
    date: formatDate_(today),
    dueDate: formatDate_(dueDate),
    businessName: settings.businessName,
    businessAddress: settings.businessAddress,
    businessEmail: settings.businessEmail,
    logoUrl: settings.logoUrl,
    clientName: clientName,
    clientEmail: clientEmail,
    clientAddress: clientAddress,
    serviceDesc: serviceDesc,
    rate: rate,
    hours: hours,
    subtotal: subtotal,
    taxRate: taxRate,
    taxAmount: taxAmount,
    total: total,
    notes: notes,
    paymentTerms: settings.paymentTerms,
  });

  // Convert to PDF
  const pdfBlob = convertDocToPdf_(docId, invoiceNumber);

  // Save PDF to Drive folder
  const pdfFile = savePdfToDrive_(pdfBlob);
  const pdfUrl = pdfFile.getUrl();

  // Clean up the temporary Google Doc
  DriveApp.getFileById(docId).setTrashed(true);

  // Email the invoice (unless test mode)
  let emailed = false;
  let status = 'Draft';
  if (!isTest && settings.autoEmail && clientEmail) {
    try {
      sendInvoiceEmail_(settings, {
        clientName: clientName,
        clientEmail: clientEmail,
        invoiceNumber: invoiceNumber,
        total: total,
        dueDate: formatDate_(dueDate),
        pdfBlob: pdfBlob,
      });
      emailed = true;
      status = 'Sent';
    } catch (e) {
      Logger.log('Email failed for ' + clientEmail + ': ' + e.message);
      status = 'Error';
    }
  }

  // Log the invoice
  logInvoice_({
    invoiceNumber: invoiceNumber,
    clientName: clientName,
    clientEmail: clientEmail,
    total: total,
    dateCreated: today,
    dueDate: dueDate,
    status: status,
    pdfUrl: pdfUrl,
  });

  // Increment invoice number
  settings.nextInvoiceNumber = parseInt(settings.nextInvoiceNumber, 10) + 1;
  saveSettings(settings);

  return {
    invoiceNumber: invoiceNumber,
    clientName: clientName,
    clientEmail: clientEmail,
    total: total,
    emailed: emailed,
  };
}

// ═══════════════════════════════════════════
// INVOICE DOCUMENT CREATION
// ═══════════════════════════════════════════

/**
 * Creates a branded Google Doc invoice.
 * @param {Object} data - All invoice data fields.
 * @return {string} The Google Doc ID.
 */
function createInvoiceDoc_(data) {
  const doc = DocumentApp.create('Invoice ' + data.invoiceNumber);
  const body = doc.getBody();

  // Page setup
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(50);
  body.setMarginRight(50);

  // ── HEADER: Business info + Invoice number ──

  // Business name
  const headerPara = body.appendParagraph(data.businessName || 'Your Business');
  headerPara.setFontSize(20);
  headerPara.setFontFamily('Roboto');
  headerPara.setBold(true);
  headerPara.setForegroundColor('#1A1A1A');

  // Business address
  if (data.businessAddress) {
    const addrPara = body.appendParagraph(data.businessAddress);
    addrPara.setFontSize(9);
    addrPara.setFontFamily('Roboto');
    addrPara.setForegroundColor('#666666');
  }

  // Business email
  if (data.businessEmail) {
    const emailPara = body.appendParagraph(data.businessEmail);
    emailPara.setFontSize(9);
    emailPara.setFontFamily('Roboto');
    emailPara.setForegroundColor('#666666');
  }

  // Spacer
  body.appendParagraph('');

  // INVOICE title bar
  const titlePara = body.appendParagraph('INVOICE');
  titlePara.setFontSize(28);
  titlePara.setFontFamily('Roboto Mono');
  titlePara.setBold(true);
  titlePara.setForegroundColor('#C9A84C');
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // Invoice details (right-aligned)
  const detailsText = 'Invoice #: ' + data.invoiceNumber +
    '\nDate: ' + data.date +
    '\nDue Date: ' + data.dueDate;
  const detailsPara = body.appendParagraph(detailsText);
  detailsPara.setFontSize(10);
  detailsPara.setFontFamily('Roboto');
  detailsPara.setForegroundColor('#333333');
  detailsPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // Horizontal line
  body.appendHorizontalRule();
  body.appendParagraph('');

  // ── BILL TO ──
  const billToPara = body.appendParagraph('BILL TO');
  billToPara.setFontSize(9);
  billToPara.setFontFamily('Roboto Mono');
  billToPara.setBold(true);
  billToPara.setForegroundColor('#C9A84C');

  const clientInfo = data.clientName +
    (data.clientAddress ? '\n' + data.clientAddress : '') +
    (data.clientEmail ? '\n' + data.clientEmail : '');
  const clientPara = body.appendParagraph(clientInfo);
  clientPara.setFontSize(10);
  clientPara.setFontFamily('Roboto');
  clientPara.setForegroundColor('#333333');

  body.appendParagraph('');

  // ── LINE ITEMS TABLE ──
  const table = body.appendTable();

  // Table header row
  const headerRow = table.appendTableRow();
  const tableHeaders = ['Description', 'Rate', 'Qty/Hrs', 'Amount'];
  const headerWidths = [250, 90, 90, 90];

  for (let i = 0; i < tableHeaders.length; i++) {
    const cell = headerRow.appendTableCell(tableHeaders[i]);
    cell.setBackgroundColor('#1A1A1A');
    const text = cell.editAsText();
    text.setFontSize(9);
    text.setFontFamily('Roboto Mono');
    text.setBold(true);
    text.setForegroundColor('#C9A84C');
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(10);
    cell.setPaddingRight(10);
  }

  // Line item row
  const itemRow = table.appendTableRow();
  const itemValues = [
    data.serviceDesc || 'Professional Services',
    '$' + data.rate.toFixed(2),
    String(data.hours),
    '$' + (data.rate * data.hours).toFixed(2),
  ];

  for (let i = 0; i < itemValues.length; i++) {
    const cell = itemRow.appendTableCell(itemValues[i]);
    cell.setBackgroundColor('#FFFFFF');
    const text = cell.editAsText();
    text.setFontSize(10);
    text.setFontFamily('Roboto');
    text.setForegroundColor('#333333');
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(10);
    cell.setPaddingRight(10);
  }

  // Remove the auto-created empty first row if present
  if (table.getNumRows() > 2) {
    table.removeRow(0);
  }

  body.appendParagraph('');

  // ── TOTALS ──
  const subtotalPara = body.appendParagraph('Subtotal: $' + data.subtotal.toFixed(2));
  subtotalPara.setFontSize(10);
  subtotalPara.setFontFamily('Roboto');
  subtotalPara.setForegroundColor('#333333');
  subtotalPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  if (data.taxRate > 0) {
    const taxPara = body.appendParagraph('Tax (' + data.taxRate + '%): $' + data.taxAmount.toFixed(2));
    taxPara.setFontSize(10);
    taxPara.setFontFamily('Roboto');
    taxPara.setForegroundColor('#333333');
    taxPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }

  body.appendHorizontalRule();

  const totalPara = body.appendParagraph('TOTAL: $' + data.total.toFixed(2));
  totalPara.setFontSize(16);
  totalPara.setFontFamily('Roboto Mono');
  totalPara.setBold(true);
  totalPara.setForegroundColor('#1A1A1A');
  totalPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph('');

  // ── NOTES ──
  if (data.notes) {
    const notesLabel = body.appendParagraph('NOTES');
    notesLabel.setFontSize(9);
    notesLabel.setFontFamily('Roboto Mono');
    notesLabel.setBold(true);
    notesLabel.setForegroundColor('#C9A84C');

    const notesPara = body.appendParagraph(data.notes);
    notesPara.setFontSize(10);
    notesPara.setFontFamily('Roboto');
    notesPara.setForegroundColor('#666666');
    notesPara.setLineSpacing(1.4);

    body.appendParagraph('');
  }

  // ── PAYMENT TERMS ──
  if (data.paymentTerms) {
    const termsLabel = body.appendParagraph('PAYMENT TERMS');
    termsLabel.setFontSize(9);
    termsLabel.setFontFamily('Roboto Mono');
    termsLabel.setBold(true);
    termsLabel.setForegroundColor('#C9A84C');

    const termsPara = body.appendParagraph(data.paymentTerms);
    termsPara.setFontSize(9);
    termsPara.setFontFamily('Roboto');
    termsPara.setForegroundColor('#666666');
    termsPara.setLineSpacing(1.4);

    body.appendParagraph('');
  }

  // ── FOOTER ──
  body.appendHorizontalRule();
  const footerPara = body.appendParagraph(BRAND.footerText);
  footerPara.setFontSize(8);
  footerPara.setFontFamily('Roboto');
  footerPara.setForegroundColor('#AAAAAA');
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();
  return doc.getId();
}

/**
 * Converts a Google Doc to PDF.
 * @param {string} docId - The Google Doc ID.
 * @param {string} invoiceNumber - Used for the file name.
 * @return {Blob} The PDF blob.
 */
function convertDocToPdf_(docId, invoiceNumber) {
  const doc = DriveApp.getFileById(docId);
  const pdfBlob = doc.getAs('application/pdf');
  pdfBlob.setName(invoiceNumber + '.pdf');
  return pdfBlob;
}

/**
 * Saves a PDF blob to the TAKScripts Invoices folder.
 * Creates the folder if it doesn't exist.
 * @param {Blob} pdfBlob - The PDF blob to save.
 * @return {File} The saved Drive file.
 */
function savePdfToDrive_(pdfBlob) {
  let folder;
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }
  return folder.createFile(pdfBlob);
}

// ═══════════════════════════════════════════
// EMAIL
// ═══════════════════════════════════════════

/**
 * Sends a professional HTML email with the invoice PDF attached.
 * @param {Object} settings - Current settings.
 * @param {Object} invoiceData - Invoice details for template variables.
 */
function sendInvoiceEmail_(settings, invoiceData) {
  // Build subject
  const subject = (settings.emailSubject || 'Invoice {{invoice_number}}')
    .replace(/\{\{invoice_number\}\}/g, invoiceData.invoiceNumber)
    .replace(/\{\{business_name\}\}/g, settings.businessName)
    .replace(/\{\{client_name\}\}/g, invoiceData.clientName);

  // Build plain-text body for fallback
  const plainBody = (settings.emailBody || '')
    .replace(/\{\{invoice_number\}\}/g, invoiceData.invoiceNumber)
    .replace(/\{\{business_name\}\}/g, settings.businessName)
    .replace(/\{\{client_name\}\}/g, invoiceData.clientName)
    .replace(/\{\{total\}\}/g, '$' + invoiceData.total.toFixed(2))
    .replace(/\{\{due_date\}\}/g, invoiceData.dueDate);

  // Build HTML email
  const htmlBody = getEmailHtml_(settings, {
    clientName: invoiceData.clientName,
    invoiceNumber: invoiceData.invoiceNumber,
    total: '$' + invoiceData.total.toFixed(2),
    dueDate: invoiceData.dueDate,
    bodyText: plainBody,
  });

  GmailApp.sendEmail(invoiceData.clientEmail, subject, plainBody, {
    htmlBody: htmlBody,
    attachments: [invoiceData.pdfBlob],
    name: settings.businessName || 'Invoice',
  });
}

/**
 * Generates the branded HTML email template.
 * @param {Object} settings - Current settings.
 * @param {Object} data - Template data.
 * @return {string} HTML string.
 */
function getEmailHtml_(settings, data) {
  const bodyLines = data.bodyText.split('\n').map(function(line) {
    return '<p style="margin: 0 0 10px; line-height: 1.6;">' + escapeHtml_(line) + '</p>';
  }).join('');

  return '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#f4f4f4;font-family:\'Segoe UI\',system-ui,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:24px 0;">' +
    '<tr><td align="center">' +
    '<table width="600" cellpadding="0" cellspacing="0" style="background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">' +
    // Header
    '<tr><td style="background:#1A1A1A;padding:24px 32px;text-align:center;">' +
    '<span style="font-size:24px;">\uD83D\uDD77</span><br>' +
    '<span style="color:#C9A84C;font-size:16px;font-weight:700;letter-spacing:1px;">' + escapeHtml_(settings.businessName || 'TAKScripts') + '</span>' +
    '</td></tr>' +
    // Invoice badge
    '<tr><td style="padding:24px 32px 0;text-align:center;">' +
    '<div style="display:inline-block;background:#1A1A1A;color:#C9A84C;padding:8px 24px;border-radius:4px;font-family:\'Roboto Mono\',monospace;font-size:14px;font-weight:700;letter-spacing:1px;">' +
    'INVOICE ' + escapeHtml_(data.invoiceNumber) +
    '</div>' +
    '</td></tr>' +
    // Amount box
    '<tr><td style="padding:20px 32px;text-align:center;">' +
    '<div style="font-size:32px;font-weight:700;color:#1A1A1A;">' + escapeHtml_(data.total) + '</div>' +
    '<div style="font-size:13px;color:#888;margin-top:4px;">Due by ' + escapeHtml_(data.dueDate) + '</div>' +
    '</td></tr>' +
    // Divider
    '<tr><td style="padding:0 32px;"><hr style="border:none;border-top:1px solid #eee;"></td></tr>' +
    // Body text
    '<tr><td style="padding:24px 32px;font-size:14px;color:#333;">' +
    bodyLines +
    '</td></tr>' +
    // Footer
    '<tr><td style="background:#f9f9f9;padding:16px 32px;text-align:center;border-top:1px solid #eee;">' +
    '<span style="font-size:12px;color:#999;">' + BRAND.footerText + '</span>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</body></html>';
}

// ═══════════════════════════════════════════
// INVOICE LOG
// ═══════════════════════════════════════════

/**
 * Logs an invoice to the Invoice Log sheet.
 * Creates the sheet if it doesn't exist.
 * @param {Object} data - Invoice log entry data.
 */
function logInvoice_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.log);
  if (!sheet) {
    sheet = createLogSheet_(ss);
  }

  const newRow = sheet.getLastRow() + 1;
  const rowData = [
    data.invoiceNumber,
    data.clientName,
    data.clientEmail,
    data.total.toFixed(2),
    formatDate_(data.dateCreated),
    formatDate_(data.dueDate),
    data.status,
    data.pdfUrl,
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  styleDataRow_(sheet, newRow, rowData.length);
  styleStatusCell_(sheet, newRow, 7, data.status);

  // Make PDF link clickable
  const linkCell = sheet.getRange(newRow, 8);
  if (data.pdfUrl) {
    linkCell.setFormula('=HYPERLINK("' + data.pdfUrl + '","View PDF")');
    linkCell.setFontColor('#C9A84C');
  }
}

/**
 * Marks a selected invoice row in the log as "Paid".
 * Triggered by selecting a row and using the menu or running directly.
 */
function markAsPaid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.log);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No invoice log found.');
    return;
  }

  const activeRange = sheet.getActiveRange();
  if (!activeRange) {
    SpreadsheetApp.getUi().alert('Select a row in the Invoice Log first.');
    return;
  }

  const row = activeRange.getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Select a data row, not the header.');
    return;
  }

  const statusCol = 7; // Status column
  const currentStatus = sheet.getRange(row, statusCol).getValue();

  if (String(currentStatus).toLowerCase() === 'paid') {
    SpreadsheetApp.getUi().alert('This invoice is already marked as Paid.');
    return;
  }

  sheet.getRange(row, statusCol).setValue('Paid');
  styleStatusCell_(sheet, row, statusCol, 'Paid');

  const invoiceNum = sheet.getRange(row, 1).getValue();
  SpreadsheetApp.getUi().alert('Invoice ' + invoiceNum + ' marked as Paid.');
}

/**
 * Adds a right-click / edit menu option to mark as paid.
 * Called via installable trigger or can be added to the menu.
 */
function onEdit(e) {
  // If user manually types "Paid" in the status column, apply styling
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.log) return;
  if (e.range.getColumn() !== 7) return; // Status column
  if (e.range.getRow() <= 1) return;

  const value = String(e.value || '').trim();
  if (['paid', 'sent', 'draft', 'error'].includes(value.toLowerCase())) {
    styleStatusCell_(sheet, e.range.getRow(), 7, value);
  }
}

// ═══════════════════════════════════════════
// TEST RUN
// ═══════════════════════════════════════════

/**
 * Generates a sample invoice without emailing.
 * Uses the first client in the sheet (or sample data if empty).
 */
function testRun() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = loadSettings();

  if (!settings.businessName) {
    ui.alert(
      'Setup Required',
      'Please configure your business details first.\n\n' +
      'Go to: \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings',
      ui.ButtonSet.OK
    );
    return;
  }

  // Use first client or create sample data
  let testData;
  const clientSheet = ss.getSheetByName(SHEET_NAMES.clients);
  if (clientSheet && clientSheet.getLastRow() > 1) {
    testData = clientSheet.getRange(2, 1, 1, CLIENT_HEADERS.length).getValues()[0];
  } else {
    testData = [
      'Test Client',
      'test@example.com',
      '123 Test Street, Sample City, ST 12345',
      'Sample Service',
      100,
      5,
      'This is a test invoice — no email was sent.',
    ];
  }

  try {
    const invoice = generateInvoiceForRow_(testData, settings, true);
    ui.alert(
      '\uD83E\uDDEA Test Run Complete',
      'Sample invoice generated:\n\n' +
      'Invoice #: ' + invoice.invoiceNumber + '\n' +
      'Client: ' + invoice.clientName + '\n' +
      'Amount: $' + invoice.total.toFixed(2) + '\n\n' +
      'No email was sent.\n' +
      'PDF saved to Drive folder: ' + FOLDER_NAME + '\n' +
      'Check the Invoice Log for details.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Test Run Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Test run error: ' + e.stack);
  }
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Formats a Date object to a readable string.
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
 * Escapes HTML special characters to prevent XSS in email templates.
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
// FIRST-RUN INITIALIZATION
// ═══════════════════════════════════════════

/**
 * Ensures all required sheets exist. Safe to call multiple times.
 * Called automatically on first menu interaction.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_NAMES.clients)) {
    createClientsSheet_(ss);
  }

  if (!ss.getSheetByName(SHEET_NAMES.log)) {
    createLogSheet_(ss);
  }

  SpreadsheetApp.getUi().alert(
    'Setup Complete',
    'Your sheets are ready:\n\n' +
    '\u2022 ' + SHEET_NAMES.clients + ' \u2014 Add your client data here\n' +
    '\u2022 ' + SHEET_NAMES.log + ' \u2014 Invoices will be tracked here\n\n' +
    'Next step: Go to \uD83D\uDD77 TAKScripts \u2192 \u2699\uFE0F Settings to configure your business details.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
    .section-title:first-child { margin-top: 0; }
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
    .field input, .field textarea, .field select {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      font-family: inherit;
      transition: border-color 0.2s, box-shadow 0.2s;
      background: white;
    }
    .field input:focus, .field textarea:focus, .field select:focus {
      outline: none;
      border-color: #C9A84C;
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea { min-height: 80px; resize: vertical; line-height: 1.5; }
    .field .help {
      font-size: 11px;
      color: #999;
      margin-top: 4px;
      line-height: 1.4;
    }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .toggle-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 0;
    }
    .toggle-row label {
      font-size: 13px;
      font-weight: 500;
      color: #333;
      text-transform: none;
      letter-spacing: 0;
    }
    .toggle {
      position: relative;
      width: 44px;
      height: 24px;
      flex-shrink: 0;
    }
    .toggle input { opacity: 0; width: 0; height: 0; }
    .toggle .slider {
      position: absolute;
      top: 0; left: 0; right: 0; bottom: 0;
      background: #ccc;
      border-radius: 24px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .toggle .slider:before {
      content: '';
      position: absolute;
      width: 18px; height: 18px;
      left: 3px; bottom: 3px;
      background: white;
      border-radius: 50%;
      transition: transform 0.2s;
    }
    .toggle input:checked + .slider { background: #C9A84C; }
    .toggle input:checked + .slider:before { transform: translateX(20px); }
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
      background: #1a1a1a;
      color: #C9A84C;
      border: 1px solid #C9A84C;
    }
    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }
    .btn-secondary {
      background: white;
      color: #666;
      border: 1px solid #ddd;
      margin-top: 8px;
    }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .btn-init {
      background: transparent;
      color: #C9A84C;
      border: 1px dashed #C9A84C;
      margin-top: 8px;
      font-size: 12px;
    }
    .btn-init:hover { background: rgba(201,168,76,0.08); }
    .status {
      text-align: center;
      padding: 8px;
      font-size: 12px;
      margin-top: 8px;
      border-radius: 6px;
      display: none;
    }
    .status.success { display: block; background: #e8f5e9; color: #2e7d32; }
    .status.error { display: block; background: #ffebee; color: #c62828; }
    .divider { border-top: 1px solid #eee; margin: 20px 0; }
    .variables {
      background: #f5f5f5;
      border-radius: 6px;
      padding: 10px 12px;
      margin-top: 8px;
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
    <div class="sub">Auto Invoice Generator \u00B7 Settings</div>
  </div>

  <div class="form">

    <!-- ── BUSINESS INFO ── -->
    <div class="section-title">Business Details</div>

    <div class="field">
      <label>Business Name</label>
      <input type="text" id="businessName" placeholder="Your Company Name" />
    </div>

    <div class="field">
      <label>Business Address</label>
      <textarea id="businessAddress" rows="2" placeholder="123 Main St, City, State ZIP"></textarea>
    </div>

    <div class="field">
      <label>Business Email</label>
      <input type="email" id="businessEmail" placeholder="billing@yourcompany.com" />
    </div>

    <div class="field">
      <label>Logo URL (optional)</label>
      <input type="url" id="logoUrl" placeholder="https://example.com/logo.png" />
      <div class="help">Public image URL for your logo on invoices.</div>
    </div>

    <!-- ── INVOICE SETTINGS ── -->
    <div class="section-title">Invoice Settings</div>

    <div class="row">
      <div class="field">
        <label>Invoice Prefix</label>
        <input type="text" id="invoicePrefix" placeholder="INV-" />
      </div>
      <div class="field">
        <label>Next Number</label>
        <input type="number" id="nextInvoiceNumber" placeholder="1001" />
      </div>
    </div>

    <div class="row">
      <div class="field">
        <label>Tax Rate (%)</label>
        <input type="number" id="taxRate" placeholder="0" step="0.01" min="0" />
      </div>
      <div class="field">
        <label>Due Days</label>
        <input type="number" id="dueDays" placeholder="30" min="1" />
      </div>
    </div>

    <div class="field">
      <label>Payment Terms</label>
      <textarea id="paymentTerms" rows="2" placeholder="Payment due within 30 days..."></textarea>
    </div>

    <!-- ── EMAIL SETTINGS ── -->
    <div class="section-title">Email Settings</div>

    <div class="toggle-row">
      <label>Auto-send invoices by email</label>
      <div class="toggle">
        <input type="checkbox" id="autoEmail" />
        <span class="slider"></span>
      </div>
    </div>

    <div class="field">
      <label>Email Subject</label>
      <input type="text" id="emailSubject" placeholder="Invoice {{invoice_number}} from {{business_name}}" />
    </div>

    <div class="field">
      <label>Email Body</label>
      <textarea id="emailBody" rows="5" placeholder="Hi {{client_name}}, please find attached..."></textarea>
      <div class="variables">
        <div class="help" style="margin-bottom: 4px;">Template variables:</div>
        <code>{{client_name}}</code> <code>{{invoice_number}}</code>
        <code>{{total}}</code> <code>{{due_date}}</code>
        <code>{{business_name}}</code>
      </div>
    </div>

    <div class="divider"></div>

    <button class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
    <button class="btn btn-init" onclick="initSheets()">Initialize Sheets</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    // Load saved settings on open
    google.script.run.withSuccessHandler(function(settings) {
      document.getElementById('businessName').value = settings.businessName || '';
      document.getElementById('businessAddress').value = settings.businessAddress || '';
      document.getElementById('businessEmail').value = settings.businessEmail || '';
      document.getElementById('logoUrl').value = settings.logoUrl || '';
      document.getElementById('taxRate').value = settings.taxRate || 0;
      document.getElementById('dueDays').value = settings.dueDays || 30;
      document.getElementById('invoicePrefix').value = settings.invoicePrefix || 'INV-';
      document.getElementById('nextInvoiceNumber').value = settings.nextInvoiceNumber || 1001;
      document.getElementById('paymentTerms').value = settings.paymentTerms || '';
      document.getElementById('emailSubject').value = settings.emailSubject || '';
      document.getElementById('emailBody').value = settings.emailBody || '';
      document.getElementById('autoEmail').checked = settings.autoEmail || false;
    }).loadSettings();

    function save() {
      var settings = {
        businessName: document.getElementById('businessName').value,
        businessAddress: document.getElementById('businessAddress').value,
        businessEmail: document.getElementById('businessEmail').value,
        logoUrl: document.getElementById('logoUrl').value,
        taxRate: parseFloat(document.getElementById('taxRate').value) || 0,
        dueDays: parseInt(document.getElementById('dueDays').value) || 30,
        invoicePrefix: document.getElementById('invoicePrefix').value || 'INV-',
        nextInvoiceNumber: parseInt(document.getElementById('nextInvoiceNumber').value) || 1001,
        paymentTerms: document.getElementById('paymentTerms').value,
        emailSubject: document.getElementById('emailSubject').value,
        emailBody: document.getElementById('emailBody').value,
        autoEmail: document.getElementById('autoEmail').checked,
      };

      var statusEl = document.getElementById('status');
      statusEl.className = 'status';
      statusEl.style.display = 'none';

      google.script.run
        .withSuccessHandler(function() {
          statusEl.textContent = '\u2713 Settings saved successfully';
          statusEl.className = 'status success';
        })
        .withFailureHandler(function(err) {
          statusEl.textContent = '\u2715 Error: ' + err.message;
          statusEl.className = 'status error';
        })
        .saveSettings(settings);
    }

    function initSheets() {
      google.script.run
        .withSuccessHandler(function() {
          var statusEl = document.getElementById('status');
          statusEl.textContent = '\u2713 Sheets initialized';
          statusEl.className = 'status success';
        })
        .withFailureHandler(function(err) {
          var statusEl = document.getElementById('status');
          statusEl.textContent = '\u2715 Error: ' + err.message;
          statusEl.className = 'status error';
        })
        .initializeSheets();
    }
  </script>
</body>
</html>`;
}
