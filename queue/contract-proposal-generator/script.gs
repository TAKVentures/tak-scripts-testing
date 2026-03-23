/**
 * Contract & Proposal Generator by TAKScripts
 * =============================================
 * Generate branded proposals and contracts from client data,
 * merge into Google Doc templates, convert to PDF, and email.
 *
 * Features:
 * - Custom "🕷 TAKScripts" menu — no script editor needed
 * - Branded settings sidebar for business info & defaults
 * - Pulls client data from a "👥 Clients" sheet
 * - Merges into Google Doc templates (proposal or contract)
 * - Creates polished PDFs saved to a Drive folder
 * - Optionally emails PDF to client with branded HTML email
 * - "📋 Document Log" sheet tracks every generated document
 * - Auto-incrementing document numbers (PROP-001, CONTRACT-001)
 * - Mark documents as Signed or Expired from the log
 * - Test Run generates a doc without emailing
 * - Fully branded: dark + gold design language
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const BRAND = {
  headerBg: '#1A1A1A',
  headerText: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  fontMono: 'Roboto Mono',
  fontRegular: 'Roboto',
  statusColors: {
    Sent:    { bg: '#E8F5E9', text: '#2E7D32' },
    Draft:   { bg: '#FFF8E1', text: '#F57F17' },
    Expired: { bg: '#FFEBEE', text: '#C62828' },
    Signed:  { bg: '#E3F2FD', text: '#1565C0' },
  },
};

const CLIENT_SHEET_NAME = '👥 Clients';
const LOG_SHEET_NAME = '📋 Document Log';
const FOLDER_NAME = '📁 TAKScripts — Proposals';
const SETTINGS_KEY = 'cpg_settings';
const DOC_COUNTER_KEY = 'cpg_doc_counter';

const CLIENT_HEADERS = [
  'Company', 'Contact Name', 'Email', 'Address',
  'Project Description', 'Scope', 'Deliverables',
  'Timeline', 'Rate/Price', 'Payment Terms',
];

const LOG_HEADERS = [
  'Doc Number', 'Client', 'Type', 'Date Created',
  'Amount', 'Status', 'PDF Link',
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
    .addSeparator()
    .addItem('📝 Generate Proposal', 'generateProposal')
    .addItem('📄 Generate Contract', 'generateContract')
    .addSeparator()
    .addItem('🧪 Test Run', 'testRun')
    .addItem('📊 View Document Log', 'viewDocumentLog')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
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
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Segoe UI', system-ui, sans-serif; padding: 20px; text-align: center;">
      <div style="font-size: 32px; margin-bottom: 8px;">🕷</div>
      <h2 style="margin: 0 0 4px; font-size: 18px;">Contract & Proposal Generator</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 · by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
        Part of the <strong>TAKScripts</strong> collection.<br>
        Pre-built Google Apps Scripts for small business.
      </p>
      <p style="margin-top: 12px; font-size: 12px; color: #888; line-height: 1.5;">
        Generate branded proposals and contracts,<br>
        convert to PDF, and email — all from a spreadsheet.
      </p>
      <p style="margin-top: 16px;">
        <a href="https://takscripts.store" target="_blank"
           style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">
          takscripts.store →
        </a>
      </p>
    </div>
  `).setWidth(300).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to the Document Log sheet.
 */
function viewDocumentLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No document log yet.\n\nGenerate a proposal or contract and it will appear here.'
    );
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
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SETTINGS_KEY, JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings with sensible defaults.
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
    defaultPaymentTerms: 'Net 30 — 50% deposit required to begin work, balance due upon completion.',
    emailSubject: 'Your {{type}} from {{business}} — {{docNumber}}',
    emailBody: 'Hi {{contactName}},\n\nPlease find your {{type}} attached. Review at your convenience and let me know if you have any questions.\n\nBest regards,\n{{business}}',
  };
  if (!raw) return defaults;
  return { ...defaults, ...JSON.parse(raw) };
}

// ═══════════════════════════════════════════
// SHEET INITIALIZATION
// ═══════════════════════════════════════════

/**
 * Ensures the Clients sheet exists with proper headers and formatting.
 */
function ensureClientSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CLIENT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CLIENT_SHEET_NAME);
    sheet.appendRow(CLIENT_HEADERS);
    styleHeaderRow_(sheet, CLIENT_HEADERS.length);

    // Set column widths
    const widths = [160, 150, 200, 200, 250, 200, 200, 120, 100, 160];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // Add sample row
    sheet.appendRow([
      'Acme Corp', 'Jane Smith', 'jane@acme.com', '123 Business Ave, Suite 100',
      'Website redesign with modern UI/UX', 'Full design and development of a responsive website',
      'Wireframes, Design mockups, Developed website, Documentation',
      '6 weeks', '$5,000', 'Net 30 — 50% deposit upfront',
    ]);

    // Style data row
    const dataRange = sheet.getRange(2, 1, 1, CLIENT_HEADERS.length);
    dataRange.setFontFamily(BRAND.fontRegular).setFontSize(10);
  }

  return sheet;
}

/**
 * Ensures the Document Log sheet exists with proper headers and formatting.
 */
function ensureLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(LOG_HEADERS);
    styleHeaderRow_(sheet, LOG_HEADERS.length);

    const widths = [140, 160, 100, 140, 100, 100, 300];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  return sheet;
}

/**
 * Applies branded header styling to a sheet.
 */
function styleHeaderRow_(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setFontFamily(BRAND.fontMono)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBackground(BRAND.headerBg)
    .setFontColor(BRAND.headerText)
    .setHorizontalAlignment('left');
  sheet.setFrozenRows(1);
}

/**
 * Applies alternating row colors and status highlighting to the log.
 */
function styleLogSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numCols = LOG_HEADERS.length;

  for (let row = 2; row <= lastRow; row++) {
    const rowRange = sheet.getRange(row, 1, 1, numCols);
    const bgColor = (row % 2 === 0) ? BRAND.lightGray : BRAND.white;
    rowRange.setBackground(bgColor).setFontFamily(BRAND.fontRegular).setFontSize(10);

    // Style status cell
    const statusCell = sheet.getRange(row, 6);
    const status = statusCell.getValue();
    if (BRAND.statusColors[status]) {
      statusCell
        .setBackground(BRAND.statusColors[status].bg)
        .setFontColor(BRAND.statusColors[status].text)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
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
 * Test Run — generates document without emailing.
 */
function testRun() {
  generateDocument_('Proposal', true);
}

/**
 * Main generation engine. Reads client data, creates doc, converts to PDF,
 * optionally emails, and logs everything.
 */
function generateDocument_(docType, isTest) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = loadSettings();

  // Validate settings
  if (!settings.businessName) {
    ui.alert('⚠️ Business name not set.\n\nOpen 🕷 TAKScripts → ⚙️ Settings to configure your business details.');
    return;
  }

  // Ensure sheets exist
  const clientSheet = ensureClientSheet();
  const logSheet = ensureLogSheet();

  // Get selected row
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== CLIENT_SHEET_NAME) {
    ui.alert('⚠️ Select a client first.\n\nGo to the "' + CLIENT_SHEET_NAME + '" sheet and click on the row of the client you want to generate a ' + docType.toLowerCase() + ' for.');
    return;
  }

  const activeRow = ss.getActiveRange().getRow();
  if (activeRow < 2) {
    ui.alert('⚠️ Select a client row.\n\nClick on a data row (not the header) in the "' + CLIENT_SHEET_NAME + '" sheet.');
    return;
  }

  const lastCol = CLIENT_HEADERS.length;
  const rowData = clientSheet.getRange(activeRow, 1, 1, lastCol).getValues()[0];

  // Map to named fields
  const client = {
    company:     rowData[0] || '',
    contactName: rowData[1] || '',
    email:       rowData[2] || '',
    address:     rowData[3] || '',
    description: rowData[4] || '',
    scope:       rowData[5] || '',
    deliverables:rowData[6] || '',
    timeline:    rowData[7] || '',
    price:       rowData[8] || '',
    paymentTerms:rowData[9] || settings.defaultPaymentTerms,
  };

  if (!client.company) {
    ui.alert('⚠️ Missing company name.\n\nThe selected row does not have a company name in column A.');
    return;
  }

  // Generate document number
  const prefix = docType === 'Proposal' ? settings.proposalPrefix : settings.contractPrefix;
  const docNumber = getNextDocNumber_(prefix);

  // Confirm with user
  const testLabel = isTest ? ' (TEST — no email will be sent)' : '';
  const confirm = ui.alert(
    'Generate ' + docType + testLabel,
    'Create ' + docType.toLowerCase() + ' ' + docNumber + ' for ' + client.company + '?\n\n' +
    'Contact: ' + client.contactName + '\n' +
    'Amount: ' + client.price + '\n' +
    (isTest ? '\nThis is a test run — no email will be sent.' : '\nPDF will be emailed to: ' + client.email),
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // Create the Google Doc
  const doc = createDocument_(docType, docNumber, client, settings);
  const docId = doc.getId();

  // Convert to PDF
  const pdfBlob = convertToPdf_(docId);
  const pdfFile = savePdfToDrive_(pdfBlob, docNumber + ' — ' + client.company);
  const pdfUrl = pdfFile.getUrl();

  // Email (unless test run)
  let status = 'Draft';
  if (!isTest && client.email) {
    sendDocumentEmail_(docType, docNumber, client, settings, pdfBlob);
    status = 'Sent';
  }

  // Log the document
  const logRow = [
    docNumber,
    client.company,
    docType,
    formatDate_(new Date()),
    client.price,
    status,
    pdfUrl,
  ];
  logSheet.appendRow(logRow);
  styleLogSheet_(logSheet);

  // Clean up the temp doc (we have the PDF now)
  DriveApp.getFileById(docId).setTrashed(true);

  // Show success
  const emoji = isTest ? '🧪' : '✅';
  const extra = isTest
    ? '\n\nThis was a test run — no email was sent.\nThe PDF was still saved to your Drive.'
    : '\n\nPDF emailed to: ' + client.email;

  ui.alert(
    emoji + ' ' + docType + ' Generated',
    docNumber + ' for ' + client.company + '\n\n' +
    'PDF saved to: ' + FOLDER_NAME + extra + '\n\n' +
    'View it in the 📊 Document Log.'
  );
}

// ═══════════════════════════════════════════
// DOCUMENT CREATION
// ═══════════════════════════════════════════

/**
 * Creates a polished Google Doc for the proposal or contract.
 */
function createDocument_(docType, docNumber, client, settings) {
  const doc = DocumentApp.create(docNumber + ' — ' + client.company);
  const body = doc.getBody();

  // Page margins
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(50);
  body.setMarginRight(50);

  // ── Company Header ──
  const headerTable = body.appendTable([
    [settings.businessName.toUpperCase()],
  ]);
  headerTable.setBorderWidth(0);
  const headerCell = headerTable.getRow(0).getCell(0);
  headerCell.setBackgroundColor(BRAND.headerBg);
  headerCell.setPaddingTop(16);
  headerCell.setPaddingBottom(16);
  headerCell.setPaddingLeft(20);
  headerCell.setPaddingRight(20);
  const headerPara = headerCell.getChild(0).asParagraph();
  headerPara.setForegroundColor(BRAND.headerText);
  headerPara.setFontFamily('Roboto Mono');
  headerPara.setFontSize(18);
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerPara.setBold(true);

  // Subheader with address and contact
  if (settings.businessAddress || settings.businessEmail || settings.businessPhone) {
    const subInfo = [
      settings.businessAddress,
      settings.businessEmail,
      settings.businessPhone,
    ].filter(Boolean).join('  |  ');

    const subPara = headerCell.appendParagraph(subInfo);
    subPara.setForegroundColor('#999999');
    subPara.setFontFamily('Roboto');
    subPara.setFontSize(9);
    subPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  body.appendParagraph('');

  // ── Document Title & Number ──
  const titleText = docType.toUpperCase();
  const titlePara = body.appendParagraph(titleText);
  titlePara.setFontFamily('Roboto Mono');
  titlePara.setFontSize(22);
  titlePara.setBold(true);
  titlePara.setForegroundColor(BRAND.headerBg);

  const numberPara = body.appendParagraph(docNumber + '  ·  ' + formatDate_(new Date()));
  numberPara.setFontFamily('Roboto');
  numberPara.setFontSize(10);
  numberPara.setForegroundColor('#888888');

  body.appendParagraph('');

  // ── Divider ──
  appendDivider_(body);

  // ── Prepared For / Client Info ──
  const preparedPara = body.appendParagraph('PREPARED FOR');
  preparedPara.setFontFamily('Roboto Mono');
  preparedPara.setFontSize(9);
  preparedPara.setForegroundColor(BRAND.headerText);
  preparedPara.setBold(true);
  preparedPara.setSpacingAfter(4);

  const clientInfoLines = [
    client.company,
    client.contactName,
    client.address,
    client.email,
  ].filter(Boolean);

  for (const line of clientInfoLines) {
    const p = body.appendParagraph(line);
    p.setFontFamily('Roboto');
    p.setFontSize(10);
    p.setForegroundColor('#333333');
    p.setSpacingAfter(2);
  }

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

  // ── Deliverables Table ──
  if (client.deliverables) {
    appendSectionHeader_(body, 'DELIVERABLES');

    const items = client.deliverables.split(',').map(d => d.trim()).filter(Boolean);
    const delivData = [['#', 'Deliverable']];
    items.forEach((item, idx) => {
      delivData.push([(idx + 1).toString(), item]);
    });

    const delivTable = body.appendTable(delivData);
    styleTable_(delivTable);

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

  // ── Pricing Table ──
  appendSectionHeader_(body, 'PRICING');

  const pricingData = [
    ['Description', 'Amount'],
    [client.description || 'Project total', client.price || 'TBD'],
  ];

  const priceTable = body.appendTable(pricingData);
  styleTable_(priceTable);

  // Bold the amount
  const amountCell = priceTable.getRow(1).getCell(1);
  amountCell.getChild(0).asParagraph().setBold(true);

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
  const termsSentences = terms.split('. ').filter(Boolean);
  for (const sentence of termsSentences) {
    const tPara = body.appendParagraph('• ' + sentence.trim() + (sentence.endsWith('.') ? '' : '.'));
    tPara.setFontFamily('Roboto');
    tPara.setFontSize(9);
    tPara.setLineSpacing(1.5);
    tPara.setForegroundColor('#555555');
  }

  body.appendParagraph('');
  appendDivider_(body);

  // ── Signature Lines ──
  appendSectionHeader_(body, 'ACCEPTANCE & SIGNATURES');

  const sigIntro = body.appendParagraph(
    'By signing below, both parties agree to the terms outlined in this ' + docType.toLowerCase() + '.'
  );
  sigIntro.setFontFamily('Roboto');
  sigIntro.setFontSize(10);
  sigIntro.setForegroundColor('#555555');

  body.appendParagraph('');

  // Signature table — two columns
  const sigData = [
    ['FOR: ' + (settings.businessName || 'Service Provider').toUpperCase(),
     'FOR: ' + client.company.toUpperCase()],
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
        para.setForegroundColor(BRAND.headerText);
      }
    }
  }

  body.appendParagraph('');

  // ── Footer ──
  const footerTable = body.appendTable([
    [settings.businessName + '  ·  takscripts.store  ·  Generated on ' + formatDate_(new Date())],
  ]);
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
 * Appends a branded section header.
 */
function appendSectionHeader_(body, text) {
  const para = body.appendParagraph(text);
  para.setFontFamily('Roboto Mono');
  para.setFontSize(10);
  para.setBold(true);
  para.setForegroundColor(BRAND.headerBg);
  para.setSpacingAfter(6);
  para.setSpacingBefore(8);
}

/**
 * Appends a visual divider line.
 */
function appendDivider_(body) {
  const divTable = body.appendTable([['']]);
  divTable.setBorderWidth(0);
  const divCell = divTable.getRow(0).getCell(0);
  divCell.setBackgroundColor(BRAND.headerText);
  divCell.setPaddingTop(0);
  divCell.setPaddingBottom(0);
  const divPara = divCell.getChild(0).asParagraph();
  divPara.setFontSize(1);
  divPara.setText('');
}

/**
 * Styles a data table with branded header and alternating rows.
 */
function styleTable_(table) {
  table.setBorderWidth(1);
  table.setBorderColor('#E0E0E0');

  // Header row
  const headerRow = table.getRow(0);
  for (let c = 0; c < headerRow.getNumCells(); c++) {
    const cell = headerRow.getCell(c);
    cell.setBackgroundColor(BRAND.headerBg);
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(10);
    cell.setPaddingRight(10);
    const para = cell.getChild(0).asParagraph();
    para.setForegroundColor(BRAND.headerText);
    para.setFontFamily('Roboto Mono');
    para.setFontSize(9);
    para.setBold(true);
  }

  // Data rows
  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const bgColor = (r % 2 === 0) ? BRAND.lightGray : BRAND.white;
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
 * Converts a Google Doc to PDF blob.
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
 * Saves PDF to the TAKScripts proposals folder. Creates folder if needed.
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
 */
function sendDocumentEmail_(docType, docNumber, client, settings, pdfBlob) {
  const subject = (settings.emailSubject || 'Your {{type}} — {{docNumber}}')
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
 * Returns branded HTML email template.
 */
function getEmailHtml_(docType, docNumber, client, settings, bodyText) {
  const bodyHtml = bodyText.replace(/\n/g, '<br>');

  return `<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin: 0; padding: 0; background: #f5f5f5; font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;">
  <div style="max-width: 600px; margin: 0 auto; background: #ffffff;">

    <!-- Header -->
    <div style="background: #1A1A1A; padding: 28px 32px; text-align: center;">
      <div style="font-size: 28px; margin-bottom: 4px;">🕷</div>
      <div style="font-size: 16px; font-weight: 700; letter-spacing: 1px;">
        <span style="color: #C9A84C;">` + (settings.businessName || 'TAKScripts').toUpperCase() + `</span>
      </div>
    </div>

    <!-- Body -->
    <div style="padding: 32px; line-height: 1.7; color: #333333; font-size: 14px;">
      ` + bodyHtml + `

      <div style="margin-top: 24px; padding: 16px; background: #F9F9F9; border-radius: 8px; border-left: 3px solid #C9A84C;">
        <div style="font-size: 12px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px;">
          Document Details
        </div>
        <div style="font-size: 13px; color: #333;">
          <strong>` + docType + `:</strong> ` + docNumber + `<br>
          <strong>Client:</strong> ` + client.company + `<br>
          <strong>Amount:</strong> ` + (client.price || 'See attached') + `
        </div>
      </div>
    </div>

    <!-- Footer -->
    <div style="background: #F5F5F5; padding: 20px 32px; text-align: center; border-top: 1px solid #eee;">
      <div style="font-size: 11px; color: #999; line-height: 1.6;">
        ` + (settings.businessName || '') + `
        ` + (settings.businessAddress ? '<br>' + settings.businessAddress : '') + `
        ` + (settings.businessPhone ? '<br>' + settings.businessPhone : '') + `
        <br><br>
        <span style="color: #C9A84C;">Powered by TAKScripts</span>
      </div>
    </div>

  </div>
</body>
</html>`;
}

// ═══════════════════════════════════════════
// DOCUMENT NUMBERING
// ═══════════════════════════════════════════

/**
 * Returns the next auto-incrementing document number.
 * Format: PREFIX001, PREFIX002, etc.
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
 * Pads a number with leading zeros.
 */
function padNumber_(num, length) {
  let str = num.toString();
  while (str.length < length) str = '0' + str;
  return str;
}

// ═══════════════════════════════════════════
// STATUS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Marks the selected document in the log as Signed.
 * Called from a custom menu or sidebar.
 */
function markAsSigned() {
  updateDocStatus_('Signed');
}

/**
 * Marks the selected document in the log as Expired.
 */
function markAsExpired() {
  updateDocStatus_('Expired');
}

/**
 * Updates the status of the selected row in the Document Log.
 */
function updateDocStatus_(newStatus) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== LOG_SHEET_NAME) {
    ui.alert('⚠️ Navigate to the "' + LOG_SHEET_NAME + '" sheet first, then select the row to update.');
    return;
  }

  const row = ss.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('⚠️ Select a document row (not the header).');
    return;
  }

  const docNumber = sheet.getRange(row, 1).getValue();
  const client = sheet.getRange(row, 2).getValue();

  const confirm = ui.alert(
    'Update Status',
    'Mark ' + docNumber + ' (' + client + ') as "' + newStatus + '"?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const statusCell = sheet.getRange(row, 6);
  statusCell.setValue(newStatus);

  // Apply status styling
  if (BRAND.statusColors[newStatus]) {
    statusCell
      .setBackground(BRAND.statusColors[newStatus].bg)
      .setFontColor(BRAND.statusColors[newStatus].text)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  ui.alert('✅ ' + docNumber + ' marked as ' + newStatus + '.');
}

// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Formats a date as "March 22, 2026".
 */
function formatDate_(date) {
  const months = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

// ═══════════════════════════════════════════
// FIRST-RUN SETUP
// ═══════════════════════════════════════════

/**
 * One-time setup function. Creates all sheets with proper formatting.
 * Safe to run multiple times — idempotent.
 */
function setupSheets() {
  ensureClientSheet();
  ensureLogSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(CLIENT_SHEET_NAME));

  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete!\n\n' +
    '1. Open ⚙️ Settings to add your business info\n' +
    '2. Add clients to the "' + CLIENT_SHEET_NAME + '" sheet\n' +
    '3. Select a client row and run 📝 Generate Proposal\n\n' +
    'Tip: Use 🧪 Test Run to preview without sending emails.'
  );
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

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
    }
    .field input:focus, .field textarea:focus {
      outline: none;
      border-color: #C9A84C;
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea {
      min-height: 80px;
      resize: vertical;
      line-height: 1.5;
    }
    .field .help {
      font-size: 11px;
      color: #999;
      margin-top: 4px;
      line-height: 1.4;
    }

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

    .btn-status {
      background: white;
      color: #333;
      border: 1px solid #ddd;
      margin-top: 6px;
      font-size: 12px;
      padding: 10px;
    }
    .btn-status:hover { border-color: #C9A84C; color: #C9A84C; }

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
    <div class="logo">🕷</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">Contract & Proposal Generator · Settings</div>
  </div>

  <div class="form">

    <!-- Business Info -->
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

    <!-- Document Numbers -->
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

    <!-- Default Terms -->
    <div class="section-title">Default Terms</div>

    <div class="field">
      <label>Terms & Conditions</label>
      <textarea id="defaultTerms" rows="4" placeholder="Payment is due within 30 days..."></textarea>
      <div class="help">Appears in every proposal and contract. Can be customized per client in the sheet.</div>
    </div>

    <div class="field">
      <label>Payment Terms</label>
      <textarea id="defaultPaymentTerms" rows="2" placeholder="Net 30 — 50% deposit required..."></textarea>
    </div>

    <!-- Email Template -->
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
        <code>{{contactName}}</code>
        <code>{{type}}</code>
        <code>{{business}}</code>
        <code>{{docNumber}}</code>
      </div>
    </div>

    <div class="divider"></div>

    <button class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div class="divider"></div>

    <div class="section-title">Quick Actions</div>
    <button class="btn btn-status" onclick="markSigned()">✅ Mark Selected Doc as Signed</button>
    <button class="btn btn-status" onclick="markExpired()">❌ Mark Selected Doc as Expired</button>
    <button class="btn btn-status" onclick="runSetup()" style="margin-top: 12px;">🔧 Run Initial Setup</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    // Load saved settings on open
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
        proposalPrefix: document.getElementById('proposalPrefix').value,
        contractPrefix: document.getElementById('contractPrefix').value,
        defaultTerms: document.getElementById('defaultTerms').value,
        defaultPaymentTerms: document.getElementById('defaultPaymentTerms').value,
        emailSubject: document.getElementById('emailSubject').value,
        emailBody: document.getElementById('emailBody').value,
      };

      showStatus('', '');
      google.script.run
        .withSuccessHandler(function() { showStatus('Settings saved successfully', 'success'); })
        .withFailureHandler(function(err) { showStatus('Error: ' + err.message, 'error'); })
        .saveSettings(settings);
    }

    function markSigned() {
      google.script.run
        .withFailureHandler(function(err) { showStatus('Error: ' + err.message, 'error'); })
        .markAsSigned();
    }

    function markExpired() {
      google.script.run
        .withFailureHandler(function(err) { showStatus('Error: ' + err.message, 'error'); })
        .markAsExpired();
    }

    function runSetup() {
      google.script.run
        .withFailureHandler(function(err) { showStatus('Error: ' + err.message, 'error'); })
        .setupSheets();
    }

    function showStatus(msg, type) {
      var el = document.getElementById('status');
      if (!msg) { el.style.display = 'none'; return; }
      el.textContent = (type === 'success' ? '✓ ' : '✕ ') + msg;
      el.className = 'status ' + type;
    }
  </script>
</body>
</html>`;
}
