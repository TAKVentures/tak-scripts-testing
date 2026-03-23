/**
 * Auto File Organizer by TAKScripts
 * ====================================
 * Automatically organizes files in a Google Drive folder by sorting
 * them into labeled subfolders based on configurable rules.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Three sorting modes: by file type, by date, or by naming pattern
 * - Emoji-prefixed subfolders for visual clarity
 * - Preview/Test Run mode — see what would move without moving anything
 * - Full activity log with timestamps and file paths
 * - Auto-schedule support (daily or weekly via triggers)
 * - Duplicate-safe — appends number if file already exists in target
 * - Idempotent — safe to run multiple times
 * - Never moves files out of the source folder tree
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

/**
 * Subfolder definitions keyed by MIME type prefix or extension group.
 * Each entry maps a category name to its emoji prefix and matching MIME types.
 */
var FILE_CATEGORIES = {
  'Documents': {
    emoji: '\uD83D\uDCC4',
    mimeTypes: [
      'application/vnd.google-apps.document',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'text/plain',
      'text/html',
      'text/csv',
      'application/rtf',
    ],
  },
  'Spreadsheets': {
    emoji: '\uD83D\uDCCA',
    mimeTypes: [
      'application/vnd.google-apps.spreadsheet',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ],
  },
  'Presentations': {
    emoji: '\uD83D\uDCBB',
    mimeTypes: [
      'application/vnd.google-apps.presentation',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ],
  },
  'PDFs': {
    emoji: '\uD83D\uDCD1',
    mimeTypes: [
      'application/pdf',
    ],
  },
  'Images': {
    emoji: '\uD83D\uDDBC',
    mimeTypes: [
      'image/jpeg',
      'image/png',
      'image/gif',
      'image/bmp',
      'image/svg+xml',
      'image/webp',
      'image/tiff',
      'application/vnd.google-apps.photo',
      'application/vnd.google-apps.drawing',
    ],
  },
  'Videos': {
    emoji: '\uD83C\uDFA5',
    mimeTypes: [
      'video/mp4',
      'video/quicktime',
      'video/x-msvideo',
      'video/webm',
      'video/x-matroska',
      'application/vnd.google-apps.video',
    ],
  },
  'Audio': {
    emoji: '\uD83C\uDFB5',
    mimeTypes: [
      'audio/mpeg',
      'audio/wav',
      'audio/ogg',
      'audio/mp4',
      'audio/x-m4a',
      'audio/flac',
      'application/vnd.google-apps.audio',
    ],
  },
  'Archives': {
    emoji: '\uD83D\uDDC4',
    mimeTypes: [
      'application/zip',
      'application/x-rar-compressed',
      'application/x-7z-compressed',
      'application/gzip',
      'application/x-tar',
    ],
  },
  'Other': {
    emoji: '\uD83D\uDCE6',
    mimeTypes: [],  // Catch-all for anything not matched above
  },
};

var DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Organizer Dashboard';
var DASHBOARD_HEADER_ROW = 3;

/**
 * Default settings for new installs.
 */
var DEFAULT_SETTINGS = {
  sourceFolderId: '',
  sourceFolderName: '',
  sortMode: 'type',         // 'type', 'date', or 'pattern'
  customRules: '',           // newline-separated regex=>FolderName rules
  scheduleEnabled: false,
  scheduleFrequency: 'daily', // 'daily' or 'weekly'
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
    .addItem('\u25B6\uFE0F Organize Now', 'organizeNow')
    .addItem('\uD83E\uDDEA Test Run (Preview)', 'testRun')
    .addSeparator()
    .addItem('\uD83D\uDCCA View Dashboard', 'viewActivityLog')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Auto File Organizer')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Auto File Organizer</h2>' +
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
 * Navigates to the Activity Log sheet, or alerts if none exists yet.
 */
function viewActivityLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No activity log yet.\n\nRun "Organize Now" or "Test Run" to generate log entries.'
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
 * @param {Object} settings - The settings object from the sidebar form.
 * @return {Object} Success indicator.
 */
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('afo_settings', JSON.stringify(settings));

  try {
    if (settings.scheduleEnabled) {
      createScheduleTrigger_(settings.scheduleFrequency);
    } else {
      removeScheduleTrigger_();
    }
  } catch (e) {
    Logger.log('Trigger update skipped: ' + e.message);
  }

  return { success: true };
}

/**
 * Load saved settings, returning defaults if none exist.
 * @return {Object} The current settings.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('afo_settings');
  if (!raw) {
    return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  }
  var saved = JSON.parse(raw);
  // Merge with defaults so new properties are always present
  var merged = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  for (var key in saved) {
    if (saved.hasOwnProperty(key)) {
      merged[key] = saved[key];
    }
  }
  return merged;
}

/**
 * Called from the sidebar to let the user pick a folder via the Drive picker.
 * Since the HTML Drive picker API is complex, we use a simple prompt approach.
 * The sidebar sends back the folder URL or ID.
 * @param {string} folderUrl - A Google Drive folder URL or folder ID.
 * @return {Object} Folder details or error.
 */
function resolveFolderFromUrl(folderUrl) {
  try {
    var folderId = folderUrl;

    // Extract ID from various URL formats
    var urlMatch = folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (urlMatch) {
      folderId = urlMatch[1];
    } else {
      // Try ?id= format
      var idMatch = folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (idMatch) {
        folderId = idMatch[1];
      }
    }

    var folder = DriveApp.getFolderById(folderId);
    return {
      success: true,
      id: folder.getId(),
      name: folder.getName(),
    };
  } catch (e) {
    return {
      success: false,
      error: 'Could not access that folder. Make sure the URL or ID is correct and you have access.',
    };
  }
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main entry point — organizes files in the source folder.
 * Called from the menu or from a scheduled trigger.
 */
function organizeNow() {
  var result = runOrganizer_(false);
  if (result.error) {
    SpreadsheetApp.getUi().alert('\u26A0\uFE0F ' + result.error);
    return;
  }

  SpreadsheetApp.getUi().alert(
    '\u2705 Organization Complete\n\n' +
    'Files moved: ' + result.moved + '\n' +
    'Files skipped (already organized): ' + result.skipped + '\n' +
    'Errors: ' + result.errors + '\n\n' +
    'Check the Activity Log for full details.'
  );
}

/**
 * Test/Preview mode — shows what would be moved without actually moving.
 */
function testRun() {
  var result = runOrganizer_(true);
  if (result.error) {
    SpreadsheetApp.getUi().alert('\u26A0\uFE0F ' + result.error);
    return;
  }

  SpreadsheetApp.getUi().alert(
    '\uD83E\uDDEA TEST RUN \u2014 Preview Only (no files moved)\n\n' +
    'Files that would be moved: ' + result.moved + '\n' +
    'Files that would be skipped: ' + result.skipped + '\n\n' +
    'Check the Activity Log for the full preview.'
  );
}

/**
 * Scheduled trigger entry point (no UI alerts).
 * Logs results to the Activity Log sheet silently.
 */
function scheduledOrganize() {
  var result = runOrganizer_(false);
  if (result.error) {
    Logger.log('Auto File Organizer scheduled run failed: ' + result.error);
    return;
  }
  Logger.log(
    'Auto File Organizer: moved=' + result.moved +
    ', skipped=' + result.skipped +
    ', errors=' + result.errors
  );
}

/**
 * Core organizer logic. Iterates files in the source folder and
 * sorts them into subfolders based on the active sort mode.
 *
 * @param {boolean} dryRun - If true, logs what would happen but moves nothing.
 * @return {Object} Summary with moved, skipped, errors counts or error message.
 */
function runOrganizer_(dryRun) {
  var settings = loadSettings();

  // Validate source folder
  if (!settings.sourceFolderId) {
    return { error: 'No source folder configured.\n\nOpen Settings from the TAKScripts menu to pick a folder.' };
  }

  var sourceFolder;
  try {
    sourceFolder = DriveApp.getFolderById(settings.sourceFolderId);
  } catch (e) {
    return { error: 'Cannot access the source folder. It may have been deleted or you lost access.\n\nOpen Settings to pick a new folder.' };
  }

  var moved = 0;
  var skipped = 0;
  var errors = 0;
  var logEntries = [];
  var mode = settings.sortMode || 'type';
  var prefix = dryRun ? '[PREVIEW] ' : '';

  // Get all files in the source folder (not recursively — only top level)
  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    try {
      var targetFolderName = getTargetFolderName_(file, mode, settings);
      if (!targetFolderName) {
        skipped++;
        continue;
      }

      // Check if file is already in the correct subfolder
      var parents = file.getParents();
      var alreadyOrganized = false;
      while (parents.hasNext()) {
        var parent = parents.next();
        if (parent.getName() === targetFolderName && parent.getId() !== sourceFolder.getId()) {
          alreadyOrganized = true;
          break;
        }
      }
      if (alreadyOrganized) {
        skipped++;
        continue;
      }

      if (!dryRun) {
        // Get or create the target subfolder
        var targetFolder = getOrCreateSubfolder_(sourceFolder, targetFolderName);
        // Handle duplicate names
        var safeName = getUniqueName_(targetFolder, file.getName());
        if (safeName !== file.getName()) {
          file.setName(safeName);
        }
        // Move the file
        file.moveTo(targetFolder);
      }

      moved++;
      logEntries.push({
        timestamp: new Date(),
        fileName: file.getName(),
        fileType: file.getMimeType(),
        from: sourceFolder.getName() + '/',
        to: sourceFolder.getName() + '/' + targetFolderName + '/',
        status: dryRun ? 'Preview' : 'Moved',
      });

    } catch (e) {
      errors++;
      logEntries.push({
        timestamp: new Date(),
        fileName: file.getName(),
        fileType: file.getMimeType(),
        from: sourceFolder.getName() + '/',
        to: 'ERROR',
        status: 'Error: ' + e.message,
      });
      Logger.log('Error processing file "' + file.getName() + '": ' + e.message);
    }
  }

  // Write log entries
  if (logEntries.length > 0) {
    writeActivityLog_(logEntries, dryRun);
  }

  return { moved: moved, skipped: skipped, errors: errors };
}

// ═══════════════════════════════════════════
// SORTING LOGIC
// ═══════════════════════════════════════════

/**
 * Determines the target subfolder name for a given file based on the sorting mode.
 *
 * @param {GoogleAppsScript.Drive.File} file - The file to categorize.
 * @param {string} mode - Sorting mode: 'type', 'date', or 'pattern'.
 * @param {Object} settings - Current settings (needed for custom pattern rules).
 * @return {string|null} The target subfolder name, or null to skip.
 */
function getTargetFolderName_(file, mode, settings) {
  switch (mode) {
    case 'type':
      return getTargetByType_(file);
    case 'date':
      return getTargetByDate_(file);
    case 'pattern':
      return getTargetByPattern_(file, settings);
    default:
      return getTargetByType_(file);
  }
}

/**
 * Sort by file type — matches MIME type to a category.
 * @param {GoogleAppsScript.Drive.File} file
 * @return {string} Subfolder name with emoji prefix.
 */
function getTargetByType_(file) {
  var mimeType = file.getMimeType();

  for (var category in FILE_CATEGORIES) {
    if (!FILE_CATEGORIES.hasOwnProperty(category)) continue;
    var cat = FILE_CATEGORIES[category];
    if (cat.mimeTypes.length === 0) continue; // Skip "Other" during matching

    for (var i = 0; i < cat.mimeTypes.length; i++) {
      if (mimeType === cat.mimeTypes[i]) {
        return cat.emoji + ' ' + category;
      }
    }
  }

  // Nothing matched — put in Other
  return FILE_CATEGORIES['Other'].emoji + ' Other';
}

/**
 * Sort by date — creates YYYY/MM folder structure based on the file's last modified date.
 * @param {GoogleAppsScript.Drive.File} file
 * @return {string} Subfolder name like "2024/03 - March".
 */
function getTargetByDate_(file) {
  var date = file.getLastUpdated();
  var year = date.getFullYear().toString();
  var monthNum = ('0' + (date.getMonth() + 1)).slice(-2);
  var monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ];
  var monthName = monthNames[date.getMonth()];
  return '\uD83D\uDCC5 ' + year + '/' + monthNum + ' - ' + monthName;
}

/**
 * Sort by naming pattern — matches filename against user-defined regex rules.
 * Rules format (one per line): regex => Folder Name
 * Example: ^INV- => Invoices
 *
 * @param {GoogleAppsScript.Drive.File} file
 * @param {Object} settings
 * @return {string|null} Subfolder name, or falls back to type-based sorting.
 */
function getTargetByPattern_(file, settings) {
  var rules = parseCustomRules_(settings.customRules || '');
  var fileName = file.getName();

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    try {
      var regex = new RegExp(rule.pattern, 'i');
      if (regex.test(fileName)) {
        return '\uD83C\uDFF7 ' + rule.folderName;
      }
    } catch (e) {
      Logger.log('Invalid regex rule: ' + rule.pattern + ' — ' + e.message);
    }
  }

  // No pattern matched — fall back to type-based sorting
  return getTargetByType_(file);
}

/**
 * Parses custom rule text into an array of {pattern, folderName} objects.
 * @param {string} rulesText - Newline-separated rules in "regex => FolderName" format.
 * @return {Array<Object>} Parsed rules.
 */
function parseCustomRules_(rulesText) {
  var lines = rulesText.split('\n');
  var rules = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line || line.indexOf('=>') === -1) continue;
    var parts = line.split('=>');
    if (parts.length >= 2) {
      rules.push({
        pattern: parts[0].trim(),
        folderName: parts.slice(1).join('=>').trim(),
      });
    }
  }
  return rules;
}

// ═══════════════════════════════════════════
// FILE & FOLDER HELPERS
// ═══════════════════════════════════════════

/**
 * Gets or creates a subfolder within the parent folder.
 * For date mode, handles nested paths like "2024/03 - March".
 *
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - The parent folder.
 * @param {string} folderName - The subfolder name (may contain / for nesting).
 * @return {GoogleAppsScript.Drive.Folder} The target subfolder.
 */
function getOrCreateSubfolder_(parentFolder, folderName) {
  // Handle nested paths (e.g., "2024/03 - March" for date sorting)
  var parts = folderName.split('/');
  var current = parentFolder;

  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;

    var existing = current.getFoldersByName(part);
    if (existing.hasNext()) {
      current = existing.next();
    } else {
      current = current.createFolder(part);
    }
  }

  return current;
}

/**
 * Returns a unique filename within the target folder.
 * If "report.pdf" exists, returns "report (1).pdf", "report (2).pdf", etc.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder - Target folder to check.
 * @param {string} name - Original filename.
 * @return {string} A name guaranteed not to conflict.
 */
function getUniqueName_(folder, name) {
  var existing = folder.getFilesByName(name);
  if (!existing.hasNext()) {
    return name; // No conflict
  }

  // Split name into base and extension
  var dotIndex = name.lastIndexOf('.');
  var baseName, extension;
  if (dotIndex > 0) {
    baseName = name.substring(0, dotIndex);
    extension = name.substring(dotIndex);
  } else {
    baseName = name;
    extension = '';
  }

  var counter = 1;
  var newName;
  do {
    newName = baseName + ' (' + counter + ')' + extension;
    counter++;
    existing = folder.getFilesByName(newName);
  } while (existing.hasNext() && counter < 100);

  return newName;
}

// ═══════════════════════════════════════════
// ACTIVITY LOG
// ═══════════════════════════════════════════

/**
 * Gets or creates the Organizer Dashboard sheet with branded stats bar.
 * Migrates from old '📋 Activity Log' sheet name if present.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The dashboard sheet.
 */
function getOrCreateDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  // Migration: rename old Activity Log sheet
  var oldSheet = ss.getSheetByName('\uD83D\uDCCB Activity Log');
  if (oldSheet && !sheet) {
    oldSheet.setName(DASHBOARD_SHEET_NAME);
    sheet = oldSheet;
    // Shift old data down by 2 rows to make room for stats bar
    var oldLastRow = sheet.getLastRow();
    if (oldLastRow >= 1) {
      sheet.insertRowsBefore(1, 2);
    }
    // Re-apply header styling on the new row 3
    var headerRange = sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, 6);
    headerRange
      .setBackground('#1A1A1A')
      .setFontColor('#C9A84C')
      .setFontFamily('Roboto Mono')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('left');
    sheet.setFrozenRows(DASHBOARD_HEADER_ROW);
    buildStatsBar_(sheet);
    refreshDashboardStats();
    return sheet;
  }

  if (sheet) {
    return sheet;
  }

  // Create fresh sheet
  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);

  buildStatsBar_(sheet);

  // Header row
  var headers = ['Timestamp', 'File Name', 'File Type', 'From', 'To', 'Status'];
  var headerRange = sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // Column widths
  sheet.setColumnWidth(1, 170);  // Timestamp
  sheet.setColumnWidth(2, 280);  // File Name
  sheet.setColumnWidth(3, 220);  // File Type
  sheet.setColumnWidth(4, 200);  // From
  sheet.setColumnWidth(5, 250);  // To
  sheet.setColumnWidth(6, 150);  // Status

  return sheet;
}

/**
 * Writes the stats bar rows 1-2 (values + labels) with brand styling.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function buildStatsBar_(sheet) {
  // Row 1: stat values (gold, large)
  var valRange = sheet.getRange(1, 1, 1, 5);
  valRange
    .setValues([['—', '—', '—', '—', '—']])
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  // Blank col 6 in stats rows
  sheet.getRange(1, 6).setBackground('#1A1A1A').setValue('');
  sheet.getRange(2, 6).setBackground('#1A1A1A').setValue('');

  // Row 2: stat labels
  var lblRange = sheet.getRange(2, 1, 1, 5);
  lblRange
    .setValues([['TOTAL MOVED', 'PREVIEWED', 'ERRORS', 'TODAY', 'LAST RUN']])
    .setBackground('#1A1A1A')
    .setFontColor('#888888')
    .setFontFamily('Roboto Mono')
    .setFontSize(8)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top');
  sheet.setRowHeight(2, 20);
}

/**
 * Reads live data and writes fresh stat values to row 1.
 * Safe to call any time.
 */
function refreshDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();

  var totalMoved = 0;
  var previewed = 0;
  var errors = 0;
  var today = 0;
  var lastRunTs = null;

  if (lastRow >= dataStartRow) {
    var data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 6).getValues();
    var now = new Date();

    for (var i = 0; i < data.length; i++) {
      var ts = data[i][0];
      var status = data[i][5];
      if (!status) continue;

      if (status === 'Moved') totalMoved++;
      else if (status === 'Preview') previewed++;
      else if (String(status).indexOf('Error') === 0) errors++;

      if (ts instanceof Date) {
        if (ts.getDate() === now.getDate() &&
            ts.getMonth() === now.getMonth() &&
            ts.getFullYear() === now.getFullYear()) {
          today++;
        }
        if (!lastRunTs || ts > lastRunTs) lastRunTs = ts;
      }
    }
  }

  var lastRunDisplay = lastRunTs
    ? Utilities.formatDate(lastRunTs, Session.getScriptTimeZone(), 'MMM d')
    : '\u2014';

  sheet.getRange(1, 1, 1, 5).setValues([[
    totalMoved || '\u2014',
    previewed || '\u2014',
    errors || '\u2014',
    today || '\u2014',
    lastRunDisplay,
  ]]);
}

/**
 * Writes file move entries to the Organizer Dashboard sheet.
 *
 * @param {Array<Object>} entries - Array of log entry objects.
 * @param {boolean} dryRun - Whether this was a preview run.
 */
function writeActivityLog_(entries, dryRun) {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;

  // Write each entry as a new row
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    var targetRow = Math.max(sheet.getLastRow() + 1, dataStartRow);
    var row = [e.timestamp, e.fileName, e.fileType, e.from, e.to, e.status];
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    // Style the status cell
    var statusCell = sheet.getRange(targetRow, 6);
    styleStatusCell_(statusCell, e.status);

    // Alternate row background
    var rowRange = sheet.getRange(targetRow, 1, 1, 6);
    if (targetRow % 2 === 0) {
      rowRange.setBackground('#F9F9F9');
    } else {
      rowRange.setBackground('#FFFFFF');
    }
    rowRange.setFontFamily('Roboto').setFontSize(10);
    // Re-apply status cell color (row bg would have overridden it)
    styleStatusCell_(sheet.getRange(targetRow, 6), e.status);
  }

  refreshDashboardStats();
}

/**
 * Applies color-coded styling to a status cell.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} cell - The cell to style.
 * @param {string} status - The status text.
 */
function styleStatusCell_(cell, status) {
  if (status === 'Moved') {
    cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
  } else if (status === 'Preview') {
    cell.setBackground('#FFF8E1').setFontColor('#F57F17');
  } else if (status.indexOf('Error') === 0) {
    cell.setBackground('#FFEBEE').setFontColor('#C62828');
  }
  cell.setFontWeight('bold');
}

// ═══════════════════════════════════════════
// SCHEDULE MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Creates a time-driven trigger for auto-organizing.
 *
 * @param {string} frequency - 'daily' or 'weekly'.
 */
function createScheduleTrigger_(frequency) {
  // Remove any existing trigger first
  removeScheduleTrigger_();

  if (frequency === 'weekly') {
    ScriptApp.newTrigger('scheduledOrganize')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
  } else {
    // Default: daily
    ScriptApp.newTrigger('scheduledOrganize')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();
  }

  Logger.log('Schedule trigger created: ' + frequency);
}

/**
 * Removes any existing schedule trigger for this script.
 */
function removeScheduleTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scheduledOrganize') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

/**
 * Returns the full HTML string for the settings sidebar.
 * Inline styles follow the TAKScripts brand guide.
 */
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
'' +
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
'    .field input, .field textarea, .field select {' +
'      width: 100%;' +
'      padding: 10px 12px;' +
'      border: 1px solid #ddd;' +
'      border-radius: 6px;' +
'      font-size: 13px;' +
'      font-family: inherit;' +
'      background: white;' +
'      transition: border-color 0.2s, box-shadow 0.2s;' +
'    }' +
'    .field input:focus, .field textarea:focus, .field select:focus {' +
'      outline: none;' +
'      border-color: #C9A84C;' +
'      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);' +
'    }' +
'    .field textarea { min-height: 100px; resize: vertical; line-height: 1.5; font-family: "Roboto Mono", monospace; font-size: 12px; }' +
'    .field .help {' +
'      font-size: 11px;' +
'      color: #999;' +
'      margin-top: 4px;' +
'      line-height: 1.4;' +
'    }' +
'' +
'    .folder-picker {' +
'      display: flex;' +
'      gap: 8px;' +
'      align-items: center;' +
'    }' +
'    .folder-picker input { flex: 1; }' +
'    .folder-picker .pick-btn {' +
'      padding: 10px 14px;' +
'      background: #1A1A1A;' +
'      color: #C9A84C;' +
'      border: 1px solid #C9A84C;' +
'      border-radius: 6px;' +
'      cursor: pointer;' +
'      font-size: 12px;' +
'      font-weight: 600;' +
'      white-space: nowrap;' +
'      transition: all 0.2s;' +
'    }' +
'    .folder-picker .pick-btn:hover { background: #C9A84C; color: #1A1A1A; }' +
'' +
'    .folder-name {' +
'      background: #f0f0f0;' +
'      padding: 8px 12px;' +
'      border-radius: 6px;' +
'      font-size: 12px;' +
'      color: #333;' +
'      margin-top: 6px;' +
'      display: none;' +
'    }' +
'    .folder-name.visible { display: block; }' +
'    .folder-name .icon { margin-right: 4px; }' +
'' +
'    .toggle-row {' +
'      display: flex;' +
'      justify-content: space-between;' +
'      align-items: center;' +
'      padding: 10px 0;' +
'    }' +
'    .toggle-row label { margin-bottom: 0; }' +
'    .toggle {' +
'      position: relative;' +
'      width: 44px;' +
'      height: 24px;' +
'    }' +
'    .toggle input { opacity: 0; width: 0; height: 0; }' +
'    .toggle .slider {' +
'      position: absolute;' +
'      cursor: pointer;' +
'      top: 0; left: 0; right: 0; bottom: 0;' +
'      background: #ccc;' +
'      border-radius: 24px;' +
'      transition: 0.3s;' +
'    }' +
'    .toggle .slider:before {' +
'      content: "";' +
'      position: absolute;' +
'      height: 18px;' +
'      width: 18px;' +
'      left: 3px;' +
'      bottom: 3px;' +
'      background: white;' +
'      border-radius: 50%;' +
'      transition: 0.3s;' +
'    }' +
'    .toggle input:checked + .slider { background: #C9A84C; }' +
'    .toggle input:checked + .slider:before { transform: translateX(20px); }' +
'' +
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
'' +
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
'    .status.warning { display: block; background: #FFF8E1; color: #F57F17; }' +
'' +
'    .divider { border-top: 1px solid #eee; margin: 20px 0; }' +
'' +
'    .section-title {' +
'      font-size: 12px;' +
'      font-weight: 700;' +
'      text-transform: uppercase;' +
'      letter-spacing: 1.5px;' +
'      color: #C9A84C;' +
'      margin-bottom: 12px;' +
'    }' +
'' +
'    .mode-cards {' +
'      display: flex;' +
'      flex-direction: column;' +
'      gap: 8px;' +
'    }' +
'    .mode-card {' +
'      display: flex;' +
'      align-items: center;' +
'      padding: 10px 12px;' +
'      border: 2px solid #eee;' +
'      border-radius: 8px;' +
'      cursor: pointer;' +
'      transition: all 0.2s;' +
'    }' +
'    .mode-card:hover { border-color: #ccc; }' +
'    .mode-card.selected { border-color: #C9A84C; background: #FFFDF5; }' +
'    .mode-card input { display: none; }' +
'    .mode-card .mode-icon { font-size: 20px; margin-right: 10px; }' +
'    .mode-card .mode-label { font-weight: 600; font-size: 12px; }' +
'    .mode-card .mode-desc { font-size: 11px; color: #888; }' +
'' +
'    .pattern-rules { display: none; }' +
'    .pattern-rules.visible { display: block; }' +
'' +
'    .footer-text {' +
'      text-align: center;' +
'      font-size: 10px;' +
'      color: #bbb;' +
'      margin-top: 20px;' +
'      padding-top: 12px;' +
'      border-top: 1px solid #eee;' +
'    }' +
'  </style>' +
'</head>' +
'<body>' +
'' +
'  <div class="header">' +
'    <div class="logo">\uD83D\uDD77</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Auto File Organizer \u00B7 Settings</div>' +
'  </div>' +
'' +
'  <div class="form">' +
'' +
'    <!-- SOURCE FOLDER -->' +
'    <div class="section-title">Source Folder</div>' +
'    <div class="field">' +
'      <label>Folder URL or ID</label>' +
'      <div class="folder-picker">' +
'        <input type="text" id="folderUrl" placeholder="Paste Google Drive folder URL or ID" />' +
'        <button class="pick-btn" onclick="resolveFolder()">Set</button>' +
'      </div>' +
'      <div id="folderName" class="folder-name">' +
'        <span class="icon">\uD83D\uDCC1</span> <span id="folderNameText">No folder selected</span>' +
'      </div>' +
'      <div class="help">Open the folder in Google Drive, copy the URL from your browser, and paste it here.</div>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <!-- SORTING MODE -->' +
'    <div class="section-title">Sorting Mode</div>' +
'    <div class="field">' +
'      <div class="mode-cards">' +
'        <label class="mode-card selected" id="modeCard_type" onclick="selectMode(\'type\')">' +
'          <input type="radio" name="sortMode" value="type" checked />' +
'          <span class="mode-icon">\uD83D\uDCC1</span>' +
'          <div>' +
'            <div class="mode-label">By File Type</div>' +
'            <div class="mode-desc">PDFs, Images, Docs, Sheets, etc.</div>' +
'          </div>' +
'        </label>' +
'        <label class="mode-card" id="modeCard_date" onclick="selectMode(\'date\')">' +
'          <input type="radio" name="sortMode" value="date" />' +
'          <span class="mode-icon">\uD83D\uDCC5</span>' +
'          <div>' +
'            <div class="mode-label">By Date</div>' +
'            <div class="mode-desc">Year/Month folders (YYYY/MM)</div>' +
'          </div>' +
'        </label>' +
'        <label class="mode-card" id="modeCard_pattern" onclick="selectMode(\'pattern\')">' +
'          <input type="radio" name="sortMode" value="pattern" />' +
'          <span class="mode-icon">\uD83C\uDFF7</span>' +
'          <div>' +
'            <div class="mode-label">By Naming Pattern</div>' +
'            <div class="mode-desc">Custom regex rules</div>' +
'          </div>' +
'        </label>' +
'      </div>' +
'    </div>' +
'' +
'    <!-- CUSTOM RULES (visible only in pattern mode) -->' +
'    <div class="field pattern-rules" id="patternRules">' +
'      <label>Custom Rules</label>' +
'      <textarea id="customRules" placeholder="^INV- => Invoices&#10;^IMG_ => Photos&#10;^DRAFT => Drafts"></textarea>' +
'      <div class="help">One rule per line. Format: <strong>regex => Folder Name</strong><br>Unmatched files fall back to file-type sorting.</div>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <!-- SCHEDULE -->' +
'    <div class="section-title">Auto Schedule</div>' +
'    <div class="field">' +
'      <div class="toggle-row">' +
'        <label>Run automatically</label>' +
'        <div class="toggle">' +
'          <input type="checkbox" id="scheduleEnabled" onchange="toggleSchedule()" />' +
'          <span class="slider"></span>' +
'        </div>' +
'      </div>' +
'    </div>' +
'' +
'    <div class="field" id="frequencyField" style="display:none;">' +
'      <label>Frequency</label>' +
'      <select id="scheduleFrequency">' +
'        <option value="daily">Daily (every morning at 9 AM)</option>' +
'        <option value="weekly">Weekly (every Monday at 9 AM)</option>' +
'      </select>' +
'    </div>' +
'' +
'    <div class="divider"></div>' +
'' +
'    <!-- ACTIONS -->' +
'    <button class="btn btn-primary" onclick="save()">\u2699\uFE0F Save Settings</button>' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
'' +
'    <div id="status" class="status"></div>' +
'' +
'    <div class="footer-text">' +
'      Powered by TAKScripts \u00B7 takscripts.store' +
'    </div>' +
'  </div>' +
'' +
'  <script>' +
'    var currentFolderId = "";' +
'    var currentFolderName = "";' +
'' +
'    // Load saved settings on open' +
'    google.script.run.withSuccessHandler(function(s) {' +
'      currentFolderId = s.sourceFolderId || "";' +
'      currentFolderName = s.sourceFolderName || "";' +
'' +
'      if (currentFolderName) {' +
'        document.getElementById("folderName").className = "folder-name visible";' +
'        document.getElementById("folderNameText").textContent = currentFolderName;' +
'        document.getElementById("folderUrl").value = currentFolderId;' +
'      }' +
'' +
'      selectMode(s.sortMode || "type");' +
'      document.getElementById("customRules").value = s.customRules || "";' +
'      document.getElementById("scheduleEnabled").checked = s.scheduleEnabled || false;' +
'      document.getElementById("scheduleFrequency").value = s.scheduleFrequency || "daily";' +
'      toggleSchedule();' +
'    }).loadSettings();' +
'' +
'    function selectMode(mode) {' +
'      var cards = document.querySelectorAll(".mode-card");' +
'      for (var i = 0; i < cards.length; i++) {' +
'        cards[i].className = "mode-card";' +
'      }' +
'      var selected = document.getElementById("modeCard_" + mode);' +
'      if (selected) {' +
'        selected.className = "mode-card selected";' +
'        selected.querySelector("input").checked = true;' +
'      }' +
'' +
'      // Show/hide pattern rules' +
'      var patternEl = document.getElementById("patternRules");' +
'      patternEl.className = mode === "pattern" ? "field pattern-rules visible" : "field pattern-rules";' +
'    }' +
'' +
'    function toggleSchedule() {' +
'      var enabled = document.getElementById("scheduleEnabled").checked;' +
'      document.getElementById("frequencyField").style.display = enabled ? "block" : "none";' +
'    }' +
'' +
'    function resolveFolder() {' +
'      var url = document.getElementById("folderUrl").value.trim();' +
'      if (!url) {' +
'        showStatus("Please enter a folder URL or ID.", "warning");' +
'        return;' +
'      }' +
'' +
'      showStatus("Looking up folder...", "warning");' +
'' +
'      google.script.run' +
'        .withSuccessHandler(function(result) {' +
'          if (result.success) {' +
'            currentFolderId = result.id;' +
'            currentFolderName = result.name;' +
'            document.getElementById("folderName").className = "folder-name visible";' +
'            document.getElementById("folderNameText").textContent = result.name;' +
'            document.getElementById("folderUrl").value = result.id;' +
'            showStatus("Folder set: " + result.name, "success");' +
'          } else {' +
'            showStatus(result.error, "error");' +
'          }' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          showStatus("Error: " + err.message, "error");' +
'        })' +
'        .resolveFolderFromUrl(url);' +
'    }' +
'' +
'    function save() {' +
'      var mode = "type";' +
'      var radios = document.querySelectorAll("input[name=sortMode]");' +
'      for (var i = 0; i < radios.length; i++) {' +
'        if (radios[i].checked) { mode = radios[i].value; break; }' +
'      }' +
'' +
'      if (!currentFolderId) {' +
'        showStatus("Please set a source folder first.", "warning");' +
'        return;' +
'      }' +
'' +
'      var settings = {' +
'        sourceFolderId: currentFolderId,' +
'        sourceFolderName: currentFolderName,' +
'        sortMode: mode,' +
'        customRules: document.getElementById("customRules").value,' +
'        scheduleEnabled: document.getElementById("scheduleEnabled").checked,' +
'        scheduleFrequency: document.getElementById("scheduleFrequency").value,' +
'      };' +
'' +
'      google.script.run' +
'        .withSuccessHandler(function() {' +
'          showStatus("Settings saved successfully!", "success");' +
'        })' +
'        .withFailureHandler(function(err) {' +
'          showStatus("Error: " + err.message, "error");' +
'        })' +
'        .saveSettings(settings);' +
'    }' +
'' +
'    function showStatus(msg, type) {' +
'      var el = document.getElementById("status");' +
'      el.textContent = (type === "success" ? "\u2713 " : type === "error" ? "\u2715 " : "\u26A0 ") + msg;' +
'      el.className = "status " + type;' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
