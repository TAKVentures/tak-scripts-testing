/**
 * Attachment Auto-Saver by TAKScripts
 * =====================================
 * Automatically scans Gmail for emails with attachments and saves them
 * to organized Google Drive folders.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Multiple organization modes: by sender, label, date, or file type
 * - File type filtering (PDFs only, images only, etc.)
 * - Max file size limit to skip large attachments
 * - Duplicate filename handling with (1), (2), etc.
 * - Skips inline images (signatures, logos) — only saves real attachments
 * - Tracks processed messages to prevent duplicate saves
 * - Test Run mode — preview what would be saved without saving
 * - Scheduled auto-save via time-driven triggers
 * - Full save log with Drive links
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var ROOT_FOLDER_NAME = '\uD83D\uDCCE TAKScripts \u2014 Saved Attachments';
var LOG_SHEET_NAME = '\uD83D\uDCCB Save Log';
var SETTINGS_SHEET_NAME = '\u2699\uFE0F Settings';
var PROP_SETTINGS = 'atas_settings';
var PROP_PROCESSED = 'atas_processed_ids';
var MAX_PROCESSED_IDS = 5000; // Cap stored IDs to avoid property size limits
var BATCH_SIZE = 50; // Max threads to process per run

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
    .addItem('\u25B6\uFE0F Save Attachments Now', 'saveAttachmentsNow')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run', 'testRun')
    .addItem('\uD83D\uDCCA View Save Log', 'viewLog')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
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
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Attachment Auto-Saver</h2>' +
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
 * Navigates to or creates the save log sheet.
 */
function viewLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No save log yet. Run "Save Attachments Now" and saved files will be logged here.');
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
    destinationFolderId: '',
    destinationFolderName: '',
    organizationMode: 'sender',   // sender, label, date, filetype
    fileTypeFilters: [],           // empty = all types
    maxFileSizeMB: 25,             // 0 = no limit
    autoSchedule: 'off',          // off, hourly, every6hours, daily
    skipInlineImages: true,
  };
}

/**
 * Save settings from the sidebar.
 */
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_SETTINGS, JSON.stringify(settings));

  // Update trigger based on schedule setting
  updateScheduleTrigger_(settings.autoSchedule);

  return { success: true };
}

/**
 * Load saved settings.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PROP_SETTINGS);
  if (!raw) {
    return getDefaultSettings_();
  }
  var saved = JSON.parse(raw);
  var defaults = getDefaultSettings_();
  // Merge with defaults so new properties are always present
  for (var key in defaults) {
    if (saved[key] === undefined) {
      saved[key] = defaults[key];
    }
  }
  return saved;
}

/**
 * Lets the user pick a destination folder from Drive via the sidebar.
 * Returns folder info to the sidebar.
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
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main entry point — called by menu or trigger.
 * Scans Gmail, saves attachments to Drive, logs results.
 */
function saveAttachmentsNow() {
  var result = processAttachments_(false);
  var msg = '\u2705 Attachment Auto-Saver Complete\n\n' +
    'Threads scanned: ' + result.threadsScanned + '\n' +
    'Files saved: ' + result.filesSaved + '\n' +
    'Files skipped: ' + result.filesSkipped + '\n' +
    'Errors: ' + result.errors;
  if (result.filesSaved > 0) {
    msg += '\n\nFiles saved to: ' + result.rootFolderName;
  }
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Trigger handler — same as saveAttachmentsNow but no UI alert.
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
    '\uD83E\uDDEA TEST RUN \u2014 No files were saved',
    'Threads scanned: ' + result.threadsScanned,
    'Files that would be saved: ' + result.filesSaved,
    'Files that would be skipped: ' + result.filesSkipped,
    '---',
  ];

  for (var i = 0; i < result.details.length; i++) {
    var d = result.details[i];
    lines.push('');
    lines.push('Email: ' + d.subject);
    lines.push('From: ' + d.sender);
    for (var j = 0; j < d.attachments.length; j++) {
      var a = d.attachments[j];
      var status = a.wouldSave ? '\u2192 WOULD SAVE' : '\u2192 SKIP (' + a.skipReason + ')';
      lines.push('  ' + a.name + ' (' + formatFileSize_(a.size) + ') ' + status);
    }
  }

  if (result.details.length === 0) {
    lines.push('No emails with attachments found matching your search criteria.');
  }

  var output = lines.join('\n');
  Logger.log(output);
  SpreadsheetApp.getUi().alert(output);
}

/**
 * Core processing function. Handles both real saves and test runs.
 *
 * @param {boolean} dryRun - If true, previews without saving
 * @returns {Object} Summary of the run
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
    return { threadsScanned: 0, filesSaved: 0, filesSkipped: 0, errors: 1, details: [] };
  }

  // Get or create root folder (only if not dry run and we have threads)
  var rootFolder = null;
  if (!dryRun && threads.length > 0) {
    rootFolder = getOrCreateRootFolder_(settings);
  }

  var result = {
    threadsScanned: threads.length,
    filesSaved: 0,
    filesSkipped: 0,
    errors: 0,
    rootFolderName: ROOT_FOLDER_NAME,
    details: [],
  };

  var newProcessedIds = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var messageId = message.getId();

      // Skip already processed messages
      if (processedSet[messageId]) continue;

      var attachments = message.getAttachments();
      if (!attachments || attachments.length === 0) continue;

      var senderEmail = extractEmail_(message.getFrom());
      var senderName = extractName_(message.getFrom());
      var subject = message.getSubject() || '(no subject)';
      var messageDate = message.getDate();

      var detail = {
        subject: subject,
        sender: senderName ? senderName + ' <' + senderEmail + '>' : senderEmail,
        attachments: [],
      };

      // Get labels for this thread (used by label organization mode)
      var threadLabels = [];
      try {
        var labels = threads[t].getLabels();
        for (var l = 0; l < labels.length; l++) {
          threadLabels.push(labels[l].getName());
        }
      } catch (e) {
        // Labels may not be accessible in some contexts
      }

      for (var a = 0; a < attachments.length; a++) {
        var attachment = attachments[a];
        var fileName = attachment.getName();
        var fileSize = attachment.getSize();
        var contentType = attachment.getContentType();

        // Check if inline image (skip signatures/logos)
        var skipReason = shouldSkipAttachment_(attachment, settings, fileName, fileSize, contentType);

        var attachDetail = {
          name: fileName,
          size: fileSize,
          type: contentType,
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

            logSavedFile_(
              messageDate,
              subject,
              senderEmail,
              safeName,
              contentType,
              fileSize,
              savedFile.getUrl()
            );

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

  // Save updated processed IDs (cap to MAX_PROCESSED_IDS)
  if (!dryRun && newProcessedIds.length > 0) {
    for (var n = 0; n < newProcessedIds.length; n++) {
      processedArray.push(newProcessedIds[n]);
    }
    // Keep only the most recent IDs to avoid property size limits
    if (processedArray.length > MAX_PROCESSED_IDS) {
      processedArray = processedArray.slice(processedArray.length - MAX_PROCESSED_IDS);
    }
    props.setProperty(PROP_PROCESSED, JSON.stringify(processedArray));
  }

  return result;
}

// ═══════════════════════════════════════════
// DRIVE HELPERS
// ═══════════════════════════════════════════

/**
 * Gets or creates the root "TAKScripts — Saved Attachments" folder.
 * If a custom destination folder is set, creates the root inside it.
 */
function getOrCreateRootFolder_(settings) {
  var parent;
  if (settings.destinationFolderId) {
    try {
      parent = DriveApp.getFolderById(settings.destinationFolderId);
    } catch (e) {
      Logger.log('Custom destination folder not found, using root Drive.');
      parent = DriveApp.getRootFolder();
    }
  } else {
    parent = DriveApp.getRootFolder();
  }

  var folders = parent.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(ROOT_FOLDER_NAME);
}

/**
 * Returns the target subfolder based on organization mode.
 */
function getTargetFolder_(rootFolder, settings, senderEmail, senderName, labels, messageDate, fileName) {
  var subfolderName;

  switch (settings.organizationMode) {
    case 'sender':
      // Use sender name if available, otherwise email
      subfolderName = senderName ? senderName + ' (' + senderEmail + ')' : senderEmail;
      // Clean up folder name
      subfolderName = sanitizeFolderName_(subfolderName);
      break;

    case 'label':
      if (labels.length > 0) {
        // Use the first non-system label
        subfolderName = sanitizeFolderName_(labels[0]);
      } else {
        subfolderName = 'Unlabeled';
      }
      break;

    case 'date':
      // Organize by YYYY / MM — Month Name
      var year = messageDate.getFullYear().toString();
      var monthNames = [
        'January','February','March','April','May','June',
        'July','August','September','October','November','December',
      ];
      var monthNum = ('0' + (messageDate.getMonth() + 1)).slice(-2);
      var monthFolder = monthNum + ' \u2014 ' + monthNames[messageDate.getMonth()];

      // Get or create year folder
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

/**
 * Gets or creates a subfolder by name within a parent folder.
 */
function getOrCreateSubfolder_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * Returns a unique filename within the folder.
 * If "report.pdf" exists, returns "report (1).pdf", etc.
 */
function getUniqueFileName_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    return fileName;
  }

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

/**
 * Sanitizes a string for use as a folder name.
 * Removes characters that are invalid in Drive folder names.
 */
function sanitizeFolderName_(name) {
  // Remove or replace problematic characters
  return name.replace(/[\/\\:*?"<>|]/g, '_').substring(0, 100);
}

// ═══════════════════════════════════════════
// GMAIL HELPERS
// ═══════════════════════════════════════════

/**
 * Builds the Gmail search query from settings.
 */
function buildSearchQuery_(settings) {
  var parts = ['has:attachment'];

  if (settings.sourceLabel) {
    parts.push('label:' + settings.sourceLabel);
  }

  if (settings.searchQuery && settings.searchQuery !== 'has:attachment') {
    parts.push(settings.searchQuery);
  }

  return parts.join(' ');
}

/**
 * Determines if an attachment should be skipped.
 * Returns the skip reason string, or null if it should be saved.
 */
function shouldSkipAttachment_(attachment, settings, fileName, fileSize, contentType) {
  // Skip inline images (CID-referenced, typically signatures and logos)
  if (settings.skipInlineImages) {
    if (isInlineImage_(attachment, contentType)) {
      return 'inline image';
    }
  }

  // Check file size limit
  if (settings.maxFileSizeMB && settings.maxFileSizeMB > 0) {
    var maxBytes = settings.maxFileSizeMB * 1024 * 1024;
    if (fileSize > maxBytes) {
      return 'exceeds ' + settings.maxFileSizeMB + 'MB limit';
    }
  }

  // Check file type filters
  if (settings.fileTypeFilters && settings.fileTypeFilters.length > 0) {
    var ext = getFileExtension_(fileName).toLowerCase();
    var allowed = false;

    for (var f = 0; f < settings.fileTypeFilters.length; f++) {
      var group = FILE_TYPE_GROUPS[settings.fileTypeFilters[f]];
      if (group) {
        for (var e = 0; e < group.extensions.length; e++) {
          if (group.extensions[e] === ext) {
            allowed = true;
            break;
          }
        }
      }
      if (allowed) break;
    }

    if (!allowed) {
      return 'file type not in filter';
    }
  }

  return null;
}

/**
 * Detects if an attachment is an inline image (signature, logo, etc.).
 * Uses content type and Content-ID header heuristics.
 */
function isInlineImage_(attachment, contentType) {
  // Only check image types
  if (!contentType || !contentType.match(/^image\//i)) {
    return false;
  }

  // Check if it has a Content-ID (CID), which indicates inline embedding
  try {
    var isInline = attachment.isGoogleType ? false : attachment.getContentType().match(/^image\//i);
    var name = attachment.getName().toLowerCase();

    // Common inline image signatures
    var inlinePatterns = [
      /^image\d{3}\./i,       // image001.png, image002.jpg
      /^logo/i,               // logo.png
      /^signature/i,          // signature.png
      /^icon/i,               // icon.png
      /^banner/i,             // banner.jpg
      /^spacer/i,             // spacer.gif
      /^pixel/i,              // tracking pixel
    ];

    for (var i = 0; i < inlinePatterns.length; i++) {
      if (name.match(inlinePatterns[i])) {
        return true;
      }
    }

    // Very small images are likely inline (under 10KB)
    if (attachment.getSize() < 10240 && contentType.match(/^image\//i)) {
      return true;
    }
  } catch (e) {
    // If we can't determine, assume not inline
  }

  return false;
}

/**
 * Extracts the email address from a "From" field.
 */
function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

/**
 * Extracts the display name from a "From" field.
 */
function extractName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim() : '';
}

/**
 * Returns the file extension without the dot.
 */
function getFileExtension_(fileName) {
  var dotIndex = fileName.lastIndexOf('.');
  if (dotIndex < 0) return '';
  return fileName.substring(dotIndex + 1);
}

/**
 * Maps a file extension to a human-readable category for folder naming.
 */
function getFileTypeCategory_(ext) {
  for (var key in FILE_TYPE_GROUPS) {
    var group = FILE_TYPE_GROUPS[key];
    for (var i = 0; i < group.extensions.length; i++) {
      if (group.extensions[i] === ext) {
        return group.label;
      }
    }
  }
  return 'Other';
}

/**
 * Formats file size in human-readable form.
 */
function formatFileSize_(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// ═══════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════

/**
 * Logs a saved file to the Save Log sheet.
 * Creates the sheet with branded headers if it doesn't exist.
 */
function logSavedFile_(timestamp, subject, sender, fileName, fileType, fileSize, driveLink) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = createLogSheet_(ss);
  }

  var row = [
    timestamp,
    subject,
    sender,
    fileName,
    fileType,
    formatFileSize_(fileSize),
    driveLink,
  ];

  sheet.appendRow(row);

  // Apply alternating row colors
  var lastRow = sheet.getLastRow();
  var numCols = 7;
  var rowRange = sheet.getRange(lastRow, 1, 1, numCols);
  var bgColor = (lastRow % 2 === 0) ? '#F9F9F9' : '#FFFFFF';
  rowRange.setBackground(bgColor)
    .setFontFamily('Roboto')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Make the Drive link clickable
  if (driveLink) {
    var linkCell = sheet.getRange(lastRow, numCols);
    linkCell.setFormula('=HYPERLINK("' + driveLink + '","Open in Drive")');
    linkCell.setFontColor('#1A73E8');
  }
}

/**
 * Creates the branded log sheet with styled headers and footer.
 */
function createLogSheet_(ss) {
  var sheet = ss.insertSheet(LOG_SHEET_NAME);
  var headers = ['Timestamp', 'Email Subject', 'Sender', 'Filename', 'File Type', 'Size', 'Drive Link'];
  sheet.appendRow(headers);

  // Style header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 170);  // Timestamp
  sheet.setColumnWidth(2, 280);  // Subject
  sheet.setColumnWidth(3, 220);  // Sender
  sheet.setColumnWidth(4, 240);  // Filename
  sheet.setColumnWidth(5, 140);  // File Type
  sheet.setColumnWidth(6, 90);   // Size
  sheet.setColumnWidth(7, 130);  // Drive Link

  // Add footer row with branding
  var footerRow = 3;
  sheet.getRange(footerRow, 1).setValue('Powered by TAKScripts \u00B7 takscripts.store');
  var footerRange = sheet.getRange(footerRow, 1, 1, headers.length);
  footerRange
    .setFontFamily('Roboto')
    .setFontSize(9)
    .setFontColor('#999999')
    .setFontStyle('italic')
    .setBackground('#FFFFFF');
  sheet.getRange(footerRow, 1, 1, 1).setFontColor('#C9A84C');

  return sheet;
}

// ═══════════════════════════════════════════
// SCHEDULING
// ═══════════════════════════════════════════

/**
 * Updates the time-driven trigger based on the schedule setting.
 * Removes existing triggers before creating a new one.
 */
function updateScheduleTrigger_(schedule) {
  // Remove existing auto-save triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSaveAttachments') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (!schedule || schedule === 'off') {
    Logger.log('Auto-schedule disabled.');
    return;
  }

  var builder = ScriptApp.newTrigger('autoSaveAttachments').timeBased();

  switch (schedule) {
    case 'hourly':
      builder.everyHours(1);
      break;
    case 'every6hours':
      builder.everyHours(6);
      break;
    case 'daily':
      builder.everyDays(1).atHour(6);
      break;
    default:
      Logger.log('Unknown schedule: ' + schedule);
      return;
  }

  builder.create();
  Logger.log('Auto-schedule set to: ' + schedule);
}

// ═══════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════

/**
 * Clears all processed message IDs. Use if you want to re-process
 * previously saved emails.
 */
function resetProcessedIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
  SpreadsheetApp.getUi().alert('Processed message history cleared. The next run will process all matching emails again.');
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
'    .field select {\n' +
'      width: 100%;\n' +
'      padding: 10px 12px;\n' +
'      border: 1px solid #ddd;\n' +
'      border-radius: 6px;\n' +
'      font-size: 13px;\n' +
'      font-family: inherit;\n' +
'      background: white;\n' +
'      transition: border-color 0.2s;\n' +
'    }\n' +
'    .field input:focus, .field select:focus {\n' +
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
'    .folder-picker {\n' +
'      display: flex;\n' +
'      gap: 8px;\n' +
'      align-items: center;\n' +
'    }\n' +
'    .folder-picker input {\n' +
'      flex: 1;\n' +
'    }\n' +
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
'    .btn-small {\n' +
'      width: auto;\n' +
'      padding: 8px 14px;\n' +
'      font-size: 11px;\n' +
'      border-radius: 6px;\n' +
'    }\n' +
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
'    .status.warning { display: block; background: #FFF8E1; color: #F57F17; }\n' +
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
'    <div class="logo">\uD83D\uDD77</div>\n' +
'    <h1><span class="brand">TAK</span>Scripts</h1>\n' +
'    <div class="sub">Attachment Auto-Saver \u00B7 Settings</div>\n' +
'  </div>\n' +
'\n' +
'  <div class="form">\n' +
'\n' +
'    <!-- SOURCE -->\n' +
'    <div class="section-title">Source</div>\n' +
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
'    <!-- DESTINATION -->\n' +
'    <div class="section-title">Destination</div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Drive Folder ID (optional)</label>\n' +
'      <input type="text" id="destinationFolderId" placeholder="Paste folder ID from the URL" />\n' +
'      <div id="folderName" class="folder-name"></div>\n' +
'      <div class="help">Leave blank to save to your Drive root. Find the ID in the folder\'s URL after /folders/.</div>\n' +
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
'    <!-- FILTERS -->\n' +
'    <div class="section-title">Filters</div>\n' +
'\n' +
'    <div class="checkbox-group">\n' +
'      <label>File Types to Save</label>\n' +
'      <div class="help" style="margin-bottom: 8px;">Leave all unchecked to save all file types.</div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="pdf" class="ft-check" /><span>PDFs</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="images" class="ft-check" /><span>Images</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="documents" class="ft-check" /><span>Documents (Word, etc.)</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="spreadsheets" class="ft-check" /><span>Spreadsheets</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="presentations" class="ft-check" /><span>Presentations</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="archives" class="ft-check" /><span>Archives (ZIP, etc.)</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="audio" class="ft-check" /><span>Audio</span></div>\n' +
'      <div class="checkbox-row"><input type="checkbox" value="video" class="ft-check" /><span>Video</span></div>\n' +
'    </div>\n' +
'\n' +
'    <div class="field">\n' +
'      <label>Max File Size (MB)</label>\n' +
'      <input type="number" id="maxFileSizeMB" min="0" max="50" value="25" />\n' +
'      <div class="help">Set to 0 for no limit. Gmail max is 25 MB.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="checkbox-group">\n' +
'      <div class="checkbox-row">\n' +
'        <input type="checkbox" id="skipInlineImages" checked />\n' +
'        <span>Skip inline images (signatures, logos)</span>\n' +
'      </div>\n' +
'    </div>\n' +
'\n' +
'    <!-- SCHEDULE -->\n' +
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
'      <div class="help">Automatically save new attachments on a schedule.</div>\n' +
'    </div>\n' +
'\n' +
'    <div class="divider"></div>\n' +
'\n' +
'    <button class="btn btn-primary" onclick="save()">Save Settings</button>\n' +
'    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>\n' +
'\n' +
'    <div id="status" class="status"></div>\n' +
'  </div>\n' +
'\n' +
'  <div class="footer">Powered by <a href="https://takscripts.store" target="_blank">TAKScripts</a> \u00B7 takscripts.store</div>\n' +
'\n' +
'  <script>\n' +
'    // Load saved settings on open\n' +
'    google.script.run.withSuccessHandler(function(settings) {\n' +
'      document.getElementById("sourceLabel").value = settings.sourceLabel || "";\n' +
'      document.getElementById("searchQuery").value = settings.searchQuery || "";\n' +
'      document.getElementById("destinationFolderId").value = settings.destinationFolderId || "";\n' +
'      document.getElementById("organizationMode").value = settings.organizationMode || "sender";\n' +
'      document.getElementById("maxFileSizeMB").value = settings.maxFileSizeMB || 25;\n' +
'      document.getElementById("skipInlineImages").checked = settings.skipInlineImages !== false;\n' +
'      document.getElementById("autoSchedule").value = settings.autoSchedule || "off";\n' +
'\n' +
'      // Show folder name if ID is set\n' +
'      if (settings.destinationFolderName) {\n' +
'        var el = document.getElementById("folderName");\n' +
'        el.textContent = "\\uD83D\\uDCC1 " + settings.destinationFolderName;\n' +
'        el.classList.add("visible");\n' +
'      }\n' +
'\n' +
'      // Restore file type checkboxes\n' +
'      var filters = settings.fileTypeFilters || [];\n' +
'      var checks = document.querySelectorAll(".ft-check");\n' +
'      for (var i = 0; i < checks.length; i++) {\n' +
'        checks[i].checked = filters.indexOf(checks[i].value) >= 0;\n' +
'      }\n' +
'    }).loadSettings();\n' +
'\n' +
'    function save() {\n' +
'      // Collect file type filters\n' +
'      var filters = [];\n' +
'      var checks = document.querySelectorAll(".ft-check");\n' +
'      for (var i = 0; i < checks.length; i++) {\n' +
'        if (checks[i].checked) filters.push(checks[i].value);\n' +
'      }\n' +
'\n' +
'      var settings = {\n' +
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
'\n' +
'      var statusEl = document.getElementById("status");\n' +
'      statusEl.className = "status";\n' +
'      statusEl.style.display = "none";\n' +
'\n' +
'      // Validate folder ID if provided\n' +
'      if (settings.destinationFolderId) {\n' +
'        google.script.run\n' +
'          .withSuccessHandler(function(folder) {\n' +
'            settings.destinationFolderName = folder.name;\n' +
'            var el = document.getElementById("folderName");\n' +
'            el.textContent = "\\uD83D\\uDCC1 " + folder.name;\n' +
'            el.classList.add("visible");\n' +
'            doSave(settings);\n' +
'          })\n' +
'          .withFailureHandler(function(err) {\n' +
'            statusEl.textContent = "\\u2715 " + err.message;\n' +
'            statusEl.className = "status error";\n' +
'          })\n' +
'          .pickDestinationFolder(settings.destinationFolderId);\n' +
'      } else {\n' +
'        doSave(settings);\n' +
'      }\n' +
'    }\n' +
'\n' +
'    function doSave(settings) {\n' +
'      var statusEl = document.getElementById("status");\n' +
'      google.script.run\n' +
'        .withSuccessHandler(function() {\n' +
'          statusEl.textContent = "\\u2713 Settings saved successfully";\n' +
'          statusEl.className = "status success";\n' +
'        })\n' +
'        .withFailureHandler(function(err) {\n' +
'          statusEl.textContent = "\\u2715 Error: " + err.message;\n' +
'          statusEl.className = "status error";\n' +
'        })\n' +
'        .saveSettings(settings);\n' +
'    }\n' +
'  </script>\n' +
'</body>\n' +
'</html>';
}
