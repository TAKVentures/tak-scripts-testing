/**
 * Attachment Auto-Saver by TAKScripts
 * =====================================
 * Automatically scans Gmail for emails with attachments and saves them
 * to organized Google Drive folders. The Sheet becomes your searchable
 * attachment index — find any file by sender, date, type, or category.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Full attachment dashboard with stats, search, and clickable Drive links
 * - Auto-categorization: Invoice, Contract, Receipt, Image, Report, etc.
 * - Multiple organization modes: by sender, label, date, or file type
 * - File type filtering (PDFs only, images only, etc.)
 * - Duplicate detection across emails
 * - Skips inline images (signatures, logos) — only saves real attachments
 * - Tracks processed messages to prevent duplicate saves
 * - Test Run mode — preview what would be saved without saving
 * - Scheduled auto-save via time-driven triggers
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var ROOT_FOLDER_NAME = '📎 TAKScripts — Saved Attachments';
var DASHBOARD_SHEET_NAME = '📊 Attachment Dashboard';
var SETTINGS_SHEET_NAME = '⚙️ Settings';
var PROP_SETTINGS = 'atas_settings';
var PROP_PROCESSED = 'atas_processed_ids';
var MAX_PROCESSED_IDS = 5000;
var BATCH_SIZE = 50;
var DASHBOARD_HEADER_ROW = 6; // Row where data headers start (rows 1-4 are stats, row 5 is spacer)

// Brand colors
var BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  medGray: '#666666',
  border: '#E0E0E0',
  successBg: '#E8F5E9', successText: '#2E7D32',
  warningBg: '#FFF8E1', warningText: '#F57F17',
  errorBg: '#FFEBEE', errorText: '#C62828',
  infoBg: '#E3F2FD', infoText: '#1565C0',
  accentRow: '#FFF8E7',
};

// File type filter groups
var FILE_TYPE_GROUPS = {
  'pdf': { label: 'PDFs', extensions: ['pdf'] },
  'images': { label: 'Images', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp', 'tiff', 'ico'] },
  'documents': { label: 'Documents', extensions: ['doc', 'docx', 'odt', 'rtf', 'txt', 'pages'] },
  'spreadsheets': { label: 'Spreadsheets', extensions: ['xls', 'xlsx', 'csv', 'ods', 'numbers'] },
  'presentations': { label: 'Presentations', extensions: ['ppt', 'pptx', 'odp', 'key'] },
  'archives': { label: 'Archives', extensions: ['zip', 'rar', '7z', 'tar', 'gz'] },
  'audio': { label: 'Audio', extensions: ['mp3', 'wav', 'aac', 'flac', 'ogg', 'm4a'] },
  'video': { label: 'Video', extensions: ['mp4', 'mov', 'avi', 'mkv', 'wmv', 'webm'] },
};

// Auto-categorization patterns
var CATEGORY_PATTERNS = {
  'Invoice': [/invoice/i, /bill/i, /factur/i, /inv[\-_\s]?\d/i],
  'Receipt': [/receipt/i, /payment\s*(confirm|proof)/i, /order\s*confirm/i, /purchase/i],
  'Contract': [/contract/i, /agreement/i, /nda/i, /terms/i, /legal/i, /signed/i],
  'Report': [/report/i, /analytics/i, /summary/i, /statement/i, /audit/i],
  'Proposal': [/proposal/i, /quote/i, /estimate/i, /bid/i, /rfp/i],
  'Resume': [/resume/i, /cv/i, /curriculum/i, /cover\s*letter/i],
  'Image': [/\.jpg$/i, /\.jpeg$/i, /\.png$/i, /\.gif$/i, /\.svg$/i, /\.webp$/i, /photo/i, /screenshot/i],
  'Presentation': [/\.pptx?$/i, /\.key$/i, /deck/i, /presentation/i, /slides/i],
  'Spreadsheet': [/\.xlsx?$/i, /\.csv$/i, /\.ods$/i, /spreadsheet/i, /budget/i, /tracker/i],
};

// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

/**
 * Creates the TAKScripts menu when the spreadsheet opens.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🕷 TAKScripts')
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('▶️ Save Attachments Now', 'saveAttachmentsNow')
    .addSeparator()
    .addItem('🧪 Test Run', 'testRun')
    .addItem('📊 View Dashboard', 'viewDashboard')
    .addItem('🔄 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Attachment Auto-Saver Settings')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">🕷</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Attachment Auto-Saver</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 · by TAK Ventures</p>' +
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
 * Navigates to the dashboard sheet.
 */
function viewDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('No dashboard yet. Run "Save Attachments Now" to start tracking attachments.');
    } catch(e) {
      Logger.log('No dashboard yet. Run "Save Attachments Now" to start tracking attachments.');
    }
    return;
  }
  ss.setActiveSheet(sheet);
}

// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Returns default settings.
 */
function getDefaultSettings_() {
  return {
    searchQuery: 'has:attachment',
    sourceLabel: '',
    senders: '',
    destinationFolderId: '',
    destinationFolderName: '',
    organizationMode: 'sender',
    fileTypeFilters: [],
    maxFileSizeMB: 25,
    autoSchedule: 'off',
    skipInlineImages: true,
  };
}

/**
 * Save settings from the sidebar.
 */
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_SETTINGS, JSON.stringify(settings));
  try {
    updateScheduleTrigger_(settings.autoSchedule);
  } catch (e) {
    Logger.log('Trigger update skipped: ' + e.message);
  }
  return { success: true };
}

/**
 * Load saved settings.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_SETTINGS);
  if (!raw) return getDefaultSettings_();
  var saved = JSON.parse(raw);
  var defaults = getDefaultSettings_();
  for (var key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}

/**
 * Validates a Drive folder ID.
 */
function pickDestinationFolder(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    return { id: folder.getId(), name: folder.getName() };
  } catch (e) {
    throw new Error('Could not access that folder. Check the ID and try again.');
  }
}

// ═══════════════════════════════════════════
// DASHBOARD — THE HEART OF THE PRODUCT
// ═══════════════════════════════════════════

/**
 * Creates or gets the dashboard sheet with branded design.
 * Row 1: Stats bar (Total Files | Storage Used | Files This Week | Top Sender | Duplicates Found)
 * Row 2: Spacer
 * Row 3: Data headers
 * Row 4+: Attachment data
 */
function getOrCreateDashboard_(ss) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
  var numCols = 10;
  var darkAlt = '#252525'; // slightly lighter than #1A1A1A for alternating stat cards

  // ═══════════════════════════════════════════
  // ROW 1 — Dashboard Title Bar
  // ═══════════════════════════════════════════
  sheet.getRange(1, 1, 1, numCols).merge();
  sheet.getRange(1, 1)
    .setValue('📊  ATTACHMENT DASHBOARD')
    .setFontFamily('Roboto Mono')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  // Gold bottom border on title
  sheet.getRange(1, 1, 1, numCols).setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ═══════════════════════════════════════════
  // ROWS 2-3 — Stat Cards (value + label pairs)
  // Each stat spans 2 columns for breathing room
  // ═══════════════════════════════════════════
  var statConfigs = [
    { value: '0', label: 'TOTAL FILES', cols: [1, 2], bg: BRAND.darkBg },
    { value: '0 KB', label: 'STORAGE USED', cols: [3, 4], bg: darkAlt },
    { value: '0', label: 'THIS WEEK', cols: [5, 6], bg: BRAND.darkBg },
    { value: '—', label: 'TOP SENDER', cols: [7, 8], bg: darkAlt },
    { value: '0', label: 'DUPLICATES', cols: [9, 10], bg: BRAND.darkBg },
  ];

  for (var s = 0; s < statConfigs.length; s++) {
    var stat = statConfigs[s];
    var col1 = stat.cols[0];
    var col2 = stat.cols[1];

    // Merge value cells (row 2)
    sheet.getRange(2, col1, 1, 2).merge();
    sheet.getRange(2, col1)
      .setValue(stat.value)
      .setFontFamily('Roboto Mono')
      .setFontSize(22)
      .setFontWeight('bold')
      .setFontColor(BRAND.gold)
      .setBackground(stat.bg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('bottom');

    // Merge label cells (row 3)
    sheet.getRange(3, col1, 1, 2).merge();
    sheet.getRange(3, col1)
      .setValue(stat.label)
      .setFontFamily('Roboto Mono')
      .setFontSize(8)
      .setFontWeight('bold')
      .setFontColor('#777777')
      .setBackground(stat.bg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('top');

    // Subtle side borders between stat cards
    if (s > 0) {
      sheet.getRange(2, col1, 2, 1).setBorder(null, true, null, null, null, null, '#333333', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  sheet.setRowHeight(2, 48);
  sheet.setRowHeight(3, 20);

  // Gold bottom border under stats section
  sheet.getRange(3, 1, 1, numCols).setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ═══════════════════════════════════════════
  // ROW 4 — Brand bar
  // ═══════════════════════════════════════════
  sheet.getRange(4, 1, 1, numCols).merge();
  sheet.getRange(4, 1)
    .setValue('🕷 TAKScripts — Attachment Auto-Saver v2.0')
    .setFontFamily('Roboto')
    .setFontSize(9)
    .setFontColor('#555555')
    .setBackground('#141414')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontStyle('italic');
  sheet.setRowHeight(4, 24);

  // ═══════════════════════════════════════════
  // ROW 5 — Spacer
  // ═══════════════════════════════════════════
  sheet.getRange(5, 1, 1, numCols).setBackground(BRAND.white);
  sheet.setRowHeight(5, 6);

  // ═══════════════════════════════════════════
  // ROW 6 — Data Headers
  // ═══════════════════════════════════════════
  var headers = [
    'Date Saved', 'Sender', 'Email Subject', 'Filename',
    'Category', 'File Type', 'Size', 'Drive Link',
    'Duplicate?', 'Organization'
  ];
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setBackground(BRAND.darkBg)
    .setFontColor(BRAND.gold)
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(DASHBOARD_HEADER_ROW, 34);

  // Bottom border on headers — gold accent
  headerRange.setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID);

  // Side borders between header cells — subtle
  headerRange.setBorder(null, null, null, null, true, null, '#333333', SpreadsheetApp.BorderStyle.SOLID);

  // Freeze all header rows (title + stats + headers)
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // ═══════════════════════════════════════════
  // Column Widths
  // ═══════════════════════════════════════════
  sheet.setColumnWidth(1, 160);  // Date
  sheet.setColumnWidth(2, 200);  // Sender
  sheet.setColumnWidth(3, 280);  // Subject
  sheet.setColumnWidth(4, 240);  // Filename
  sheet.setColumnWidth(5, 110);  // Category
  sheet.setColumnWidth(6, 120);  // File Type
  sheet.setColumnWidth(7, 80);   // Size
  sheet.setColumnWidth(8, 120);  // Drive Link
  sheet.setColumnWidth(9, 100);  // Duplicate?
  sheet.setColumnWidth(10, 140); // Organization

  // ═══════════════════════════════════════════
  // Footer
  // ═══════════════════════════════════════════
  var footerRow = DASHBOARD_HEADER_ROW + 2;
  sheet.getRange(footerRow, 1, 1, numCols).merge();
  sheet.getRange(footerRow, 1)
    .setValue('Powered by TAKScripts · takscripts.store')
    .setFontFamily('Roboto')
    .setFontSize(9)
    .setFontColor('#CCCCCC')
    .setFontStyle('italic')
    .setBackground(BRAND.white)
    .setHorizontalAlignment('center');

  // Move dashboard to first position
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  return sheet;
}

/**
 * Refreshes the dashboard stats bar (row 1) from the data.
 */
function refreshDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert('No dashboard yet. Run "Save Attachments Now" first.');
    } catch(e) {
      Logger.log('No dashboard yet. Run "Save Attachments Now" first.');
    }
    return;
  }

  var lastRow = sheet.getLastRow();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  if (lastRow < dataStartRow) {
    // Reset all stat values (row 2, merged cells at columns 1,3,5,7,9)
    sheet.getRange(2, 1).setValue('0');
    sheet.getRange(2, 3).setValue('0 KB');
    sheet.getRange(2, 5).setValue('0');
    sheet.getRange(2, 7).setValue('—');
    sheet.getRange(2, 9).setValue('0');
    return;
  }

  var numDataRows = lastRow - dataStartRow + 1;
  var data = sheet.getRange(dataStartRow, 1, numDataRows, 10).getValues();

  var totalFiles = 0;
  var totalSizeBytes = 0;
  var filesThisWeek = 0;
  var senderCounts = {};
  var duplicates = 0;
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || row[0] === '') continue;

    totalFiles++;

    var sizeStr = String(row[6]);
    totalSizeBytes += parseSizeToBytes_(sizeStr);

    var dateVal = row[0];
    if (dateVal instanceof Date && dateVal >= weekAgo) {
      filesThisWeek++;
    }

    var sender = String(row[1]);
    if (sender) {
      senderCounts[sender] = (senderCounts[sender] || 0) + 1;
    }

    if (String(row[8]).toLowerCase() === 'yes') {
      duplicates++;
    }
  }

  var topSender = '—';
  var topCount = 0;
  for (var s in senderCounts) {
    if (senderCounts[s] > topCount) {
      topCount = senderCounts[s];
      topSender = s;
    }
  }
  if (topSender.length > 20) topSender = topSender.substring(0, 17) + '...';

  // Update stat values (row 2, merged cells at columns 1,3,5,7,9)
  sheet.getRange(2, 1).setValue(String(totalFiles));
  sheet.getRange(2, 3).setValue(formatFileSize_(totalSizeBytes));
  sheet.getRange(2, 5).setValue(String(filesThisWeek));
  sheet.getRange(2, 7).setValue(topSender);
  sheet.getRange(2, 9).setValue(String(duplicates));
}

/**
 * Parses a human-readable size string back to bytes.
 */
function parseSizeToBytes_(sizeStr) {
  if (!sizeStr) return 0;
  var match = String(sizeStr).match(/([\d.]+)\s*(B|KB|MB|GB)/i);
  if (!match) return 0;
  var num = parseFloat(match[1]);
  var unit = match[2].toUpperCase();
  switch (unit) {
    case 'B': return num;
    case 'KB': return num * 1024;
    case 'MB': return num * 1024 * 1024;
    case 'GB': return num * 1024 * 1024 * 1024;
    default: return 0;
  }
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main entry point — called by menu or trigger.
 */
function saveAttachmentsNow() {
  var result = processAttachments_(false);
  var msg = '✅ Attachment Auto-Saver Complete\n\n' +
    'Threads scanned: ' + result.threadsScanned + '\n' +
    'Files saved: ' + result.filesSaved + '\n' +
    'Files skipped: ' + result.filesSkipped + '\n' +
    'Duplicates found: ' + result.duplicatesFound + '\n' +
    'Errors: ' + result.errors;
  if (result.filesSaved > 0) {
    msg += '\n\nFiles saved to: ' + result.rootFolderName;
    msg += '\nDashboard updated with ' + result.filesSaved + ' new entries.';
  }
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(msg);
  }
}

/**
 * Trigger handler — same logic, no UI alert.
 */
function autoSaveAttachments() {
  processAttachments_(false);
}

/**
 * Test run — shows what would be saved without saving.
 */
function testRun() {
  var result = processAttachments_(true);

  var lines = [
    '🧪 TEST RUN — No files were saved',
    'Threads scanned: ' + result.threadsScanned,
    'Files that would be saved: ' + result.filesSaved,
    'Files that would be skipped: ' + result.filesSkipped,
    'Potential duplicates: ' + result.duplicatesFound,
    '---',
  ];

  for (var i = 0; i < result.details.length; i++) {
    var d = result.details[i];
    lines.push('');
    lines.push('Email: ' + d.subject);
    lines.push('From: ' + d.sender);
    for (var j = 0; j < d.attachments.length; j++) {
      var a = d.attachments[j];
      var status = a.wouldSave ? '→ WOULD SAVE [' + a.category + ']' : '→ SKIP (' + a.skipReason + ')';
      if (a.isDuplicate) status += ' ⚠️ DUPLICATE';
      lines.push('  ' + a.name + ' (' + formatFileSize_(a.size) + ') ' + status);
    }
  }

  if (result.details.length === 0) {
    lines.push('No emails with attachments found matching your search criteria.');
  }

  try {
    SpreadsheetApp.getUi().alert(lines.join('\n'));
  } catch(e) {
    Logger.log(lines.join('\n'));
  }
}

/**
 * Core processing function.
 */
function processAttachments_(dryRun) {
  var settings = loadSettings();
  var props = PropertiesService.getScriptProperties();

  // Load processed message IDs
  var processedRaw = props.getProperty(PROP_PROCESSED) || '[]';
  var processedArray = JSON.parse(processedRaw);
  var processedSet = {};
  for (var p = 0; p < processedArray.length; p++) {
    processedSet[processedArray[p]] = true;
  }

  // Build search query
  var query = buildSearchQuery_(settings);
  Logger.log('Search query: ' + query);

  var threads;
  try {
    threads = GmailApp.search(query, 0, BATCH_SIZE);
  } catch (e) {
    Logger.log('Gmail search error: ' + e.message);
    return { threadsScanned: 0, filesSaved: 0, filesSkipped: 0, duplicatesFound: 0, errors: 1, details: [] };
  }

  // Get or create root folder and dashboard
  var rootFolder = null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = null;

  if (!dryRun && threads.length > 0) {
    rootFolder = getOrCreateRootFolder_(settings);
    dashboard = getOrCreateDashboard_(ss);
  }

  // Load existing filenames for duplicate detection
  var existingFiles = {};
  if (dashboard) {
    var lastRow = dashboard.getLastRow();
    if (lastRow > DASHBOARD_HEADER_ROW) {
      var existingData = dashboard.getRange(DASHBOARD_HEADER_ROW + 1, 4, lastRow - DASHBOARD_HEADER_ROW, 1).getValues();
      for (var e = 0; e < existingData.length; e++) {
        var fn = String(existingData[e][0]).toLowerCase();
        if (fn) existingFiles[fn] = true;
      }
    }
  }

  var result = {
    threadsScanned: threads.length,
    filesSaved: 0,
    filesSkipped: 0,
    duplicatesFound: 0,
    errors: 0,
    rootFolderName: ROOT_FOLDER_NAME,
    details: [],
  };

  var newProcessedIds = [];
  var newRows = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var messageId = message.getId();

      if (processedSet[messageId]) continue;

      var attachments = message.getAttachments();
      if (!attachments || attachments.length === 0) continue;

      var senderEmail = extractEmail_(message.getFrom());
      var senderName = extractName_(message.getFrom());
      var subject = message.getSubject() || '(no subject)';
      var messageDate = message.getDate();
      var senderDisplay = senderName ? senderName : senderEmail;

      var detail = {
        subject: subject,
        sender: senderDisplay,
        attachments: [],
      };

      var threadLabels = [];
      try {
        var labels = threads[t].getLabels();
        for (var l = 0; l < labels.length; l++) {
          threadLabels.push(labels[l].getName());
        }
      } catch (e) { /* Labels may not be accessible */ }

      for (var a = 0; a < attachments.length; a++) {
        var attachment = attachments[a];
        var fileName = attachment.getName();
        var fileSize = attachment.getSize();
        var contentType = attachment.getContentType();

        var skipReason = shouldSkipAttachment_(attachment, settings, fileName, fileSize, contentType);

        // Auto-categorize
        var category = categorizeAttachment_(fileName, subject, contentType);

        // Duplicate detection
        var isDuplicate = existingFiles[fileName.toLowerCase()] || false;
        if (isDuplicate) result.duplicatesFound++;

        var attachDetail = {
          name: fileName,
          size: fileSize,
          type: contentType,
          category: category,
          isDuplicate: isDuplicate,
          wouldSave: !skipReason,
          skipReason: skipReason || '',
        };
        detail.attachments.push(attachDetail);

        if (skipReason) {
          result.filesSkipped++;
          continue;
        }

        if (!dryRun) {
          try {
            var targetFolder = getTargetFolder_(rootFolder, settings, senderEmail, senderName, threadLabels, messageDate, fileName);
            var safeName = getUniqueFileName_(targetFolder, fileName);
            var blob = attachment.copyBlob();
            var savedFile = targetFolder.createFile(blob.setName(safeName));

            // Determine organization path
            var orgPath = getOrganizationPath_(settings, senderEmail, senderName, threadLabels, messageDate, fileName);

            // Build row for dashboard
            newRows.push({
              date: messageDate,
              sender: senderDisplay,
              subject: subject,
              filename: safeName,
              category: category,
              fileType: getFileTypeLabel_(contentType, fileName),
              size: formatFileSize_(fileSize),
              sizeBytes: fileSize,
              driveLink: savedFile.getUrl(),
              isDuplicate: isDuplicate ? 'Yes' : 'No',
              orgPath: orgPath,
            });

            existingFiles[fileName.toLowerCase()] = true;
            result.filesSaved++;
          } catch (e) {
            Logger.log('Error saving ' + fileName + ': ' + e.message);
            result.errors++;
          }
        } else {
          result.filesSaved++;
        }
      }

      if (detail.attachments.length > 0) {
        result.details.push(detail);
      }

      if (!dryRun) {
        newProcessedIds.push(messageId);
      }
    }
  }

  // Write rows to dashboard
  if (!dryRun && newRows.length > 0) {
    writeToDashboard_(dashboard, newRows);
    refreshDashboardStats();
  }

  // Save updated processed IDs
  if (!dryRun && newProcessedIds.length > 0) {
    for (var n = 0; n < newProcessedIds.length; n++) {
      processedArray.push(newProcessedIds[n]);
    }
    if (processedArray.length > MAX_PROCESSED_IDS) {
      processedArray = processedArray.slice(processedArray.length - MAX_PROCESSED_IDS);
    }
    props.setProperty(PROP_PROCESSED, JSON.stringify(processedArray));
  }

  return result;
}

// ═══════════════════════════════════════════
// DASHBOARD WRITING
// ═══════════════════════════════════════════

/**
 * Writes new rows to the dashboard with proper styling.
 */
function writeToDashboard_(sheet, rows) {
  // Remove footer row if it exists (we'll re-add it after)
  removeFooter_(sheet);

  var startRow = sheet.getLastRow() + 1;
  if (startRow <= DASHBOARD_HEADER_ROW) startRow = DASHBOARD_HEADER_ROW + 1;

  var numCols = 10;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowNum = startRow + i;

    var rowData = [
      r.date,
      r.sender,
      r.subject,
      r.filename,
      r.category,
      r.fileType,
      r.size,
      '', // Drive link (set as hyperlink formula below)
      r.isDuplicate,
      r.orgPath,
    ];

    sheet.getRange(rowNum, 1, 1, numCols).setValues([rowData]);

    // Style the row
    var rowRange = sheet.getRange(rowNum, 1, 1, numCols);
    var bgColor = ((rowNum - DASHBOARD_HEADER_ROW) % 2 === 0) ? BRAND.lightGray : BRAND.white;
    rowRange
      .setFontFamily('Roboto')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setBackground(bgColor);

    // Table borders — subtle gray grid
    rowRange.setBorder(null, null, true, null, true, null, BRAND.border, SpreadsheetApp.BorderStyle.SOLID);

    // Date formatting
    sheet.getRange(rowNum, 1).setNumberFormat('yyyy-mm-dd hh:mm');

    // Category cell — color coded
    var catCell = sheet.getRange(rowNum, 5);
    applyCategoryStyle_(catCell, r.category);

    // Drive link as clickable hyperlink
    if (r.driveLink) {
      sheet.getRange(rowNum, 8)
        .setFormula('=HYPERLINK("' + r.driveLink + '","📂 Open")')
        .setFontColor('#1A73E8')
        .setHorizontalAlignment('center');
    }

    // Duplicate cell — highlight if yes
    var dupCell = sheet.getRange(rowNum, 9);
    dupCell.setHorizontalAlignment('center');
    if (r.isDuplicate === 'Yes') {
      dupCell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText).setFontWeight('bold');
    }

    // Size — right align
    sheet.getRange(rowNum, 7).setHorizontalAlignment('right');
  }

  // Re-add footer
  addFooter_(sheet, startRow + rows.length);
}

/**
 * Applies category-specific styling to a cell.
 */
function applyCategoryStyle_(cell, category) {
  cell.setHorizontalAlignment('center').setFontWeight('bold').setFontSize(9);

  switch (category) {
    case 'Invoice':
    case 'Receipt':
      cell.setBackground(BRAND.successBg).setFontColor(BRAND.successText);
      break;
    case 'Contract':
    case 'Proposal':
      cell.setBackground(BRAND.infoBg).setFontColor(BRAND.infoText);
      break;
    case 'Report':
    case 'Spreadsheet':
      cell.setBackground(BRAND.accentRow).setFontColor(BRAND.warningText);
      break;
    case 'Resume':
      cell.setBackground('#F3E5F5').setFontColor('#7B1FA2');
      break;
    case 'Image':
    case 'Presentation':
      cell.setBackground('#E8EAF6').setFontColor('#283593');
      break;
    default: // General
      cell.setBackground('#F5F5F5').setFontColor('#616161');
  }
}

/**
 * Removes the footer row.
 */
function removeFooter_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= DASHBOARD_HEADER_ROW) return;

  for (var r = lastRow; r > DASHBOARD_HEADER_ROW; r--) {
    var val = sheet.getRange(r, 1).getValue();
    if (String(val).indexOf('Powered by') === 0) {
      sheet.deleteRow(r);
      break;
    }
  }
}

/**
 * Adds the branded footer row.
 */
function addFooter_(sheet, footerRow) {
  sheet.getRange(footerRow, 1, 1, 10).merge();
  sheet.getRange(footerRow, 1)
    .setValue('Powered by TAKScripts · takscripts.store')
    .setFontFamily('Roboto')
    .setFontSize(9)
    .setFontColor('#CCCCCC')
    .setFontStyle('italic')
    .setBackground(BRAND.white)
    .setHorizontalAlignment('center');
}

// ═══════════════════════════════════════════
// AUTO-CATEGORIZATION
// ═══════════════════════════════════════════

/**
 * Categorizes an attachment based on filename, subject, and content type.
 */
function categorizeAttachment_(fileName, subject, contentType) {
  var testString = fileName + ' ' + subject;

  for (var category in CATEGORY_PATTERNS) {
    var patterns = CATEGORY_PATTERNS[category];
    for (var i = 0; i < patterns.length; i++) {
      if (patterns[i].test(testString)) {
        return category;
      }
    }
  }

  // Fall back to content type
  if (contentType) {
    if (contentType.match(/^image\//i)) return 'Image';
    if (contentType.match(/pdf/i)) return 'Document';
    if (contentType.match(/spreadsheet|excel|csv/i)) return 'Spreadsheet';
    if (contentType.match(/presentation|powerpoint/i)) return 'Presentation';
    if (contentType.match(/video\//i)) return 'Video';
    if (contentType.match(/audio\//i)) return 'Audio';
  }

  return 'General';
}

/**
 * Returns a human-readable file type label.
 */
function getFileTypeLabel_(contentType, fileName) {
  var ext = getFileExtension_(fileName).toUpperCase();
  if (ext) return ext;
  if (!contentType) return 'Unknown';

  if (contentType.match(/pdf/i)) return 'PDF';
  if (contentType.match(/image\//i)) return 'Image';
  if (contentType.match(/video\//i)) return 'Video';
  if (contentType.match(/audio\//i)) return 'Audio';
  if (contentType.match(/zip|archive|compressed/i)) return 'Archive';
  return 'File';
}

/**
 * Returns the organization path string for display in the dashboard.
 */
function getOrganizationPath_(settings, senderEmail, senderName, labels, messageDate, fileName) {
  switch (settings.organizationMode) {
    case 'sender':
      return '📁 ' + (senderName || senderEmail);
    case 'label':
      return '🏷 ' + (labels.length > 0 ? labels[0] : 'Unlabeled');
    case 'date':
      var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return '📅 ' + messageDate.getFullYear() + '/' + monthNames[messageDate.getMonth()];
    case 'filetype':
      var ext = getFileExtension_(fileName).toLowerCase();
      return '📄 ' + getFileTypeCategory_(ext);
    default:
      return '📁 Attachments';
  }
}

// ═══════════════════════════════════════════
// DRIVE HELPERS
// ═══════════════════════════════════════════

/**
 * Gets or creates the root folder.
 */
function getOrCreateRootFolder_(settings) {
  var parent;
  if (settings.destinationFolderId) {
    try {
      parent = DriveApp.getFolderById(settings.destinationFolderId);
    } catch (e) {
      parent = DriveApp.getRootFolder();
    }
  } else {
    parent = DriveApp.getRootFolder();
  }

  var folders = parent.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(ROOT_FOLDER_NAME);
}

/**
 * Returns the target subfolder based on organization mode.
 */
function getTargetFolder_(rootFolder, settings, senderEmail, senderName, labels, messageDate, fileName) {
  var subfolderName;

  switch (settings.organizationMode) {
    case 'sender':
      subfolderName = senderName ? senderName + ' (' + senderEmail + ')' : senderEmail;
      subfolderName = sanitizeFolderName_(subfolderName);
      break;
    case 'label':
      subfolderName = (labels.length > 0) ? sanitizeFolderName_(labels[0]) : 'Unlabeled';
      break;
    case 'date':
      var year = messageDate.getFullYear().toString();
      var monthNames = ['January','February','March','April','May','June',
        'July','August','September','October','November','December'];
      var monthNum = ('0' + (messageDate.getMonth() + 1)).slice(-2);
      var monthFolder = monthNum + ' — ' + monthNames[messageDate.getMonth()];
      var yearFolder = getOrCreateSubfolder_(rootFolder, year);
      return getOrCreateSubfolder_(yearFolder, monthFolder);
    case 'filetype':
      var ext = getFileExtension_(fileName).toLowerCase();
      subfolderName = getFileTypeCategory_(ext);
      break;
    default:
      subfolderName = 'Attachments';
  }

  return getOrCreateSubfolder_(rootFolder, subfolderName);
}

function getOrCreateSubfolder_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(name);
}

function getUniqueFileName_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return fileName;

  var baseName = fileName;
  var ext = '';
  var dotIndex = fileName.lastIndexOf('.');
  if (dotIndex > 0) {
    baseName = fileName.substring(0, dotIndex);
    ext = fileName.substring(dotIndex);
  }

  var counter = 1;
  var candidate;
  do {
    candidate = baseName + ' (' + counter + ')' + ext;
    counter++;
  } while (folder.getFilesByName(candidate).hasNext());

  return candidate;
}

function sanitizeFolderName_(name) {
  return name.replace(/[\/\\:*?"<>|]/g, '_').substring(0, 100);
}

// ═══════════════════════════════════════════
// GMAIL HELPERS
// ═══════════════════════════════════════════

function buildSearchQuery_(settings) {
  var parts = ['has:attachment'];

  // Multiple senders — joined with OR into a group: (from:a OR from:b)
  if (settings.senders && settings.senders.trim()) {
    var senderList = settings.senders.split(',')
      .map(function(s) { return s.trim(); })
      .filter(function(s) { return s.length > 0; });
    if (senderList.length === 1) {
      parts.push('from:' + senderList[0]);
    } else if (senderList.length > 1) {
      parts.push('(' + senderList.map(function(s) { return 'from:' + s; }).join(' OR ') + ')');
    }
  }

  if (settings.sourceLabel) parts.push('label:' + settings.sourceLabel);
  if (settings.searchQuery && settings.searchQuery !== 'has:attachment') {
    parts.push(settings.searchQuery);
  }
  return parts.join(' ');
}

function shouldSkipAttachment_(attachment, settings, fileName, fileSize, contentType) {
  if (settings.skipInlineImages && isInlineImage_(attachment, contentType)) {
    return 'inline image';
  }
  if (settings.maxFileSizeMB && settings.maxFileSizeMB > 0) {
    var maxBytes = settings.maxFileSizeMB * 1024 * 1024;
    if (fileSize > maxBytes) return 'exceeds ' + settings.maxFileSizeMB + 'MB limit';
  }
  if (settings.fileTypeFilters && settings.fileTypeFilters.length > 0) {
    var ext = getFileExtension_(fileName).toLowerCase();
    var allowed = false;
    for (var f = 0; f < settings.fileTypeFilters.length; f++) {
      var group = FILE_TYPE_GROUPS[settings.fileTypeFilters[f]];
      if (group) {
        for (var e = 0; e < group.extensions.length; e++) {
          if (group.extensions[e] === ext) { allowed = true; break; }
        }
      }
      if (allowed) break;
    }
    if (!allowed) return 'file type not in filter';
  }
  return null;
}

function isInlineImage_(attachment, contentType) {
  if (!contentType || !contentType.match(/^image\//i)) return false;
  try {
    var name = attachment.getName().toLowerCase();
    var inlinePatterns = [
      /^image\d{3}\./i, /^logo/i, /^signature/i, /^icon/i,
      /^banner/i, /^spacer/i, /^pixel/i,
    ];
    for (var i = 0; i < inlinePatterns.length; i++) {
      if (name.match(inlinePatterns[i])) return true;
    }
    if (attachment.getSize() < 10240 && contentType.match(/^image\//i)) return true;
  } catch (e) { /* If we can't determine, assume not inline */ }
  return false;
}

function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

function extractName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim() : '';
}

function getFileExtension_(fileName) {
  var dotIndex = fileName.lastIndexOf('.');
  if (dotIndex < 0) return '';
  return fileName.substring(dotIndex + 1);
}

function getFileTypeCategory_(ext) {
  for (var key in FILE_TYPE_GROUPS) {
    var group = FILE_TYPE_GROUPS[key];
    for (var i = 0; i < group.extensions.length; i++) {
      if (group.extensions[i] === ext) return group.label;
    }
  }
  return 'Other';
}

function formatFileSize_(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  return (bytes / (1024 * 1024 * 1024)).toFixed(1) + ' GB';
}

// ═══════════════════════════════════════════
// SCHEDULING
// ═══════════════════════════════════════════

function updateScheduleTrigger_(schedule) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSaveAttachments') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  if (!schedule || schedule === 'off') return;

  var builder = ScriptApp.newTrigger('autoSaveAttachments').timeBased();
  switch (schedule) {
    case 'hourly': builder.everyHours(1); break;
    case 'every6hours': builder.everyHours(6); break;
    case 'daily': builder.everyDays(1).atHour(6); break;
    default: return;
  }
  builder.create();
}

// ═══════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════

/**
 * Clears processed message history so all emails get re-processed.
 */
function resetProcessedIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
  try {
    SpreadsheetApp.getUi().alert('Processed message history cleared. The next run will process all matching emails again.');
  } catch(e) {
    Logger.log('Processed message history cleared. The next run will process all matching emails again.');
  }
}

/**
 * Initial setup — creates the dashboard sheet.
 */
function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateDashboard_(ss);
  try {
    SpreadsheetApp.getUi().alert('✅ Dashboard created! Use the 🕷 TAKScripts menu to configure settings and start saving attachments.');
  } catch(e) {
    Logger.log('✅ Dashboard created! Use the 🕷 TAKScripts menu to configure settings and start saving attachments.');
  }
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

function getSettingsHtml() {
  return '<!DOCTYPE html>\n' +
'<html>\n' +
'<head>\n' +
'  <style>\n' +
'    * { box-sizing: border-box; margin: 0; padding: 0; }\n' +
'    body {\n' +
'      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;\n' +
'      background: #FAFAFA;\n' +
'      color: #1a1a1a;\n' +
'      font-size: 13px;\n' +
'    }\n' +
'    .header {\n' +
'      background: #1A1A1A;\n' +
'      color: white;\n' +
'      padding: 20px 16px;\n' +
'      text-align: center;\n' +
'    }\n' +
'    .header .logo { font-size: 24px; margin-bottom: 4px; }\n' +
'    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }\n' +
'    .header .brand { color: #C9A84C; }\n' +
'    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }\n' +
'\n' +
'    .form { padding: 16px; }\n' +
'    .section-title {\n' +
'      font-size: 12px;\n' +
'      font-weight: 700;\n' +
'      text-transform: uppercase;\n' +
'      letter-spacing: 1.5px;\n' +
'      color: #C9A84C;\n' +
'      margin: 20px 0 12px;\n' +
'      padding-bottom: 6px;\n' +
'      border-bottom: 1px solid #eee;\n' +
'    }\n' +
'    .section-title:first-child { margin-top: 0; }\n' +
'\n' +
'    .field { margin-bottom: 16px; }\n' +
'    .field label {\n' +
'      display: block;\n' +
'      font-size: 11px;\n' +
'      font-weight: 600;\n' +
'      text-transform: uppercase;\n' +
'      letter-spacing: 1px;\n' +
'      color: #666;\n' +
'      margin-bottom: 6px;\n' +
'    }\n' +
'    .field input[type="text"],\n' +
'    .field input[type="number"],\n' +
'    .field select,\n' +
'    .field textarea {\n' +
'      width: 100%;\n' +
'      padding: 10px 12px;\n' +
'      border: 1px solid #ddd;\n' +
'      border-radius: 6px;\n' +
'      font-size: 13px;\n' +
'      font-family: inherit;\n' +
'      background: white;\n' +
'      box-sizing: border-box;\n' +
'      transition: border-color 0.2s;\n' +
'    }\n' +
'    .field input:focus, .field select:focus, .field textarea:focus {\n' +
'      outline: none;\n' +
'      border-color: #C9A84C;\n' +
'      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);\n' +
'    }\n' +
'    .field .help {\n' +
'      font-size: 11px;\n' +
'      color: #999;\n' +
'      margin-top: 4px;\n' +
'      line-height: 1.4;\n' +
'    }\n' +
'\n' +
'    .checkbox-group { margin-bottom: 16px; }\n' +
'    .checkbox-group label {\n' +
'      display: block;\n' +
'      font-size: 11px;\n' +
'      font-weight: 600;\n' +
'      text-transform: uppercase;\n' +
'      letter-spacing: 1px;\n' +
'      color: #666;\n' +
'      margin-bottom: 8px;\n' +
'    }\n' +
'    .checkbox-row {\n' +
'      display: flex;\n' +
'      align-items: center;\n' +
'      padding: 6px 0;\n' +
'    }\n' +
'    .checkbox-row input[type="checkbox"] {\n' +
'      margin-right: 8px;\n' +
'      accent-color: #C9A84C;\n' +
'      width: 16px;\n' +
'      height: 16px;\n' +
'    }\n' +
'    .checkbox-row span {\n' +
'      font-size: 13px;\n' +
'      color: #333;\n' +
'    }\n' +
'\n' +
'    .folder-name {\n' +
'      font-size: 12px;\n' +
'      color: #333;\n' +
'      background: #f0f0f0;\n' +
'      padding: 4px 8px;\n' +
'      border-radius: 4px;\n' +
'      margin-top: 4px;\n' +
'      display: none;\n' +
'    }\n' +
'    .folder-name.visible { display: block; }\n' +
'\n' +
'    .btn {\n' +
'      width: 100%;\n' +
'      padding: 12px;\n' +
'      border: none;\n' +
'      border-radius: 8px;\n' +
'      font-size: 13px;\n' +
'      font-weight: 600;\n' +
'      cursor: pointer;\n' +
'      transition: all 0.2s;\n' +
'      letter-spacing: 0.5px;\n' +
'    }\n' +
'    .btn-primary {\n' +
'      background: #1A1A1A;\n' +
'      color: #C9A84C;\n' +
'      border: 1px solid #C9A84C;\n' +
'    }\n' +
'    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }\n' +
'    .btn-secondary {\n' +
'      background: white;\n' +
'      color: #666;\n' +
'      border: 1px solid #ddd;\n' +
'      margin-top: 8px;\n' +
'    }\n' +
'    .btn-secondary:hover { border-color: #999; color: #333; }\n' +
'\n' +
'    .status {\n' +
'      text-align: center;\n' +
'      padding: 8px;\n' +
'      font-size: 12px;\n' +
'      margin-top: 8px;\n' +
'      border-radius: 6px;\n' +
'      display: none;\n' +
'    }\n' +
'    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }\n' +
'    .status.error { display: block; background: #FFEBEE; color: #C62828; }\n' +
'\n' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }\n' +
'\n' +
'    .footer {\n' +
'      text-align: center;\n' +
'      padding: 12px 16px;\n' +
'      font-size: 10px;\n' +
'      color: #bbb;\n' +
'      border-top: 1px solid #eee;\n' +
'    }\n' +
'    .footer a { color: #C9A84C; text-decoration: none; }\n' +
'  </style>\n' +
'</head>\n' +
'<body>\n' +
'  <div class="header">\n' +
'    <div class="logo">🕷</div>\n' +
'    <h1><span class="brand">TAK</span>Scripts</h1>\n' +
'    <div class="sub">Attachment Auto-Saver · Settings</div>\n' +
'  </div>\n' +
'\n' +
'  <div class="form">\n' +
'\n' +
'    <div class="section-title">Source</div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Senders (optional)</label>\n' +
'      <textarea id="senders" rows="4" placeholder="accounting@acme.com, invoices@supplier.com, orders@shop.com" style="resize: vertical;"></textarea>\n' +
'      <div class="help">Comma-separated. Only save attachments from these senders. Leave blank for all senders.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Gmail Label (optional)</label>\n' +
'      <input type="text" id="sourceLabel" placeholder="e.g. Invoices, Receipts" />\n' +
'      <div class="help">Only scan emails with this label. Leave blank for all.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Additional Search Query</label>\n' +
'      <input type="text" id="searchQuery" placeholder="e.g. from:accounting@company.com" />\n' +
'      <div class="help">Gmail search syntax. "has:attachment" is always included.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="section-title">Destination</div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Drive Folder ID (optional)</label>\n' +
'      <input type="text" id="destinationFolderId" placeholder="Paste folder ID from the URL" />\n' +
'      <div id="folderName" class="folder-name"></div>\n' +
'      <div class="help">Leave blank to save to Drive root.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Organization Mode</label>\n' +
'      <select id="organizationMode">\n' +
'        <option value="sender">By Sender (folder per sender)</option>\n' +
'        <option value="label">By Gmail Label</option>\n' +
'        <option value="date">By Date (YYYY / MM)</option>\n' +
'        <option value="filetype">By File Type</option>\n' +
'      </select>\n' +
'    </div>\n' +
'\n' +
'    <div class="section-title">Filters</div>\n' +
'\n' +
'    <div class="checkbox-group">\n' +
'      <label>File Types to Save</label>\n' +
'      <div class="help" style="margin-bottom: 8px;">Leave all unchecked to save all file types.</div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="pdf" class="ft-check" /><span>PDFs</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="images" class="ft-check" /><span>Images</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="documents" class="ft-check" /><span>Documents</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="spreadsheets" class="ft-check" /><span>Spreadsheets</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="presentations" class="ft-check" /><span>Presentations</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="archives" class="ft-check" /><span>Archives</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="audio" class="ft-check" /><span>Audio</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="video" class="ft-check" /><span>Video</span></div>\n' +
'    </div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Max File Size (MB)</label>\n' +
'      <input type="number" id="maxFileSizeMB" min="0" max="50" value="25" />\n' +
'      <div class="help">Set to 0 for no limit.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="checkbox-group">\n' +
'      <div class="checkbox-row">\n' +
'        <input type="checkbox" id="skipInlineImages" checked />\n' +
'        <span>Skip inline images (signatures, logos)</span>\n' +
'      </div>\n' +
'    </div>\n' +
'\n' +
'    <div class="section-title">Schedule</div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Auto-Save Schedule</label>\n' +
'      <select id="autoSchedule">\n' +
'        <option value="off">Off (manual only)</option>\n' +
'        <option value="hourly">Every Hour</option>\n' +
'        <option value="every6hours">Every 6 Hours</option>\n' +
'        <option value="daily">Daily (6 AM)</option>\n' +
'      </select>\n' +
'    </div>\n' +
'\n' +
'    <div id="status" class="status"></div>\n' +
'\n' +
'    <div class="divider"></div>\n' +
'\n' +
'    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>\n' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>\n' +
'\n' +
'  </div>\n' +
'\n' +
'  <div class="footer">Powered by <a href="https://takscripts.store" target="_blank">TAKScripts</a> · takscripts.store</div>\n' +
'\n' +
'  <script>\n' +
'    google.script.run.withSuccessHandler(function(settings) {\n' +
'      document.getElementById("senders").value = settings.senders || "";\n' +
'      document.getElementById("sourceLabel").value = settings.sourceLabel || "";\n' +
'      document.getElementById("searchQuery").value = settings.searchQuery || "";\n' +
'      document.getElementById("destinationFolderId").value = settings.destinationFolderId || "";\n' +
'      document.getElementById("organizationMode").value = settings.organizationMode || "sender";\n' +
'      document.getElementById("maxFileSizeMB").value = settings.maxFileSizeMB || 25;\n' +
'      document.getElementById("skipInlineImages").checked = settings.skipInlineImages !== false;\n' +
'      document.getElementById("autoSchedule").value = settings.autoSchedule || "off";\n' +
'      if (settings.destinationFolderName) {\n' +
'        var el = document.getElementById("folderName");\n' +
'        el.textContent = "📁 " + settings.destinationFolderName;\n' +
'        el.classList.add("visible");\n' +
'      }\n' +
'      var filters = settings.fileTypeFilters || [];\n' +
'      var checks = document.querySelectorAll(".ft-check");\n' +
'      for (var i = 0; i < checks.length; i++) {\n' +
'        checks[i].checked = filters.indexOf(checks[i].value) >= 0;\n' +
'      }\n' +
'    }).loadSettings();\n' +
'\n' +
'    function save() {\n' +
'      var filters = [];\n' +
'      var checks = document.querySelectorAll(".ft-check");\n' +
'      for (var i = 0; i < checks.length; i++) {\n' +
'        if (checks[i].checked) filters.push(checks[i].value);\n' +
'      }\n' +
'      var settings = {\n' +
'        senders: document.getElementById("senders").value.trim(),\n' +
'        sourceLabel: document.getElementById("sourceLabel").value.trim(),\n' +
'        searchQuery: document.getElementById("searchQuery").value.trim(),\n' +
'        destinationFolderId: document.getElementById("destinationFolderId").value.trim(),\n' +
'        destinationFolderName: "",\n' +
'        organizationMode: document.getElementById("organizationMode").value,\n' +
'        fileTypeFilters: filters,\n' +
'        maxFileSizeMB: parseInt(document.getElementById("maxFileSizeMB").value) || 0,\n' +
'        skipInlineImages: document.getElementById("skipInlineImages").checked,\n' +
'        autoSchedule: document.getElementById("autoSchedule").value,\n' +
'      };\n' +
'      var statusEl = document.getElementById("status");\n' +
'      statusEl.className = "status";\n' +
'      statusEl.style.display = "none";\n' +
'      if (settings.destinationFolderId) {\n' +
'        google.script.run\n' +
'          .withSuccessHandler(function(folder) {\n' +
'            settings.destinationFolderName = folder.name;\n' +
'            var el = document.getElementById("folderName");\n' +
'            el.textContent = "📁 " + folder.name;\n' +
'            el.classList.add("visible");\n' +
'            doSave(settings);\n' +
'          })\n' +
'          .withFailureHandler(function(err) {\n' +
'            statusEl.textContent = "✕ " + err.message;\n' +
'            statusEl.className = "status error";\n' +
'          })\n' +
'          .pickDestinationFolder(settings.destinationFolderId);\n' +
'      } else {\n' +
'        doSave(settings);\n' +
'      }\n' +
'    }\n' +
'    function doSave(settings) {\n' +
'      var statusEl = document.getElementById("status");\n' +
'      var saveBtn = document.getElementById("saveBtn");\n' +
'      saveBtn.disabled = true;\n' +
'      saveBtn.textContent = "Saving…";\n' +
'      google.script.run\n' +
'        .withSuccessHandler(function() {\n' +
'          statusEl.textContent = "✓ Settings saved successfully";\n' +
'          statusEl.className = "status success";\n' +
'          saveBtn.textContent = "✓ Saved!";\n' +
'          setTimeout(function() {\n' +
'            saveBtn.textContent = "Save Settings";\n' +
'            saveBtn.disabled = false;\n' +
'          }, 2500);\n' +
'        })\n' +
'        .withFailureHandler(function(err) {\n' +
'          statusEl.textContent = "✕ Error: " + err.message;\n' +
'          statusEl.className = "status error";\n' +
'          saveBtn.textContent = "Save Settings";\n' +
'          saveBtn.disabled = false;\n' +
'        })\n' +
'        .saveSettings(settings);\n' +
'    }\n' +
'  </script>\n' +
'</body>\n' +
'</html>';
}
