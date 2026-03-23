/**
 * Follow-Up Nudger by TAKScripts
 * ================================
 * Automatically tracks sent emails that haven't received a reply
 * and optionally sends polite follow-up nudges.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Scans sent mail for emails with no reply after X days (default 3)
 * - Tracks pending follow-ups in a branded "Follow-Up Tracker" sheet
 * - Custom follow-up message template with {{name}}, {{subject}}, {{days}} variables
 * - Exclude specific domains from tracking
 * - Configurable max follow-ups per thread
 * - Auto-send follow-up emails (optional) or email digest notification
 * - Test Run mode — preview what would be flagged without sending anything
 * - Runs on schedule via time-driven trigger (every hour)
 * - Idempotent — tracks thread IDs already processed
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var CONFIG = {
  SHEET_NAME: '\uD83D\uDCEC Follow-Up Tracker',
  LOG_SHEET_NAME: '\uD83D\uDCDD Follow-Up Log',
  TRIGGER_FUNCTION: 'runFollowUpCheck',
  TRIGGER_INTERVAL_HOURS: 1,
  SCAN_DAYS_BACK: 14,
  HEADERS: [
    'Date Sent', 'To', 'Subject', 'Days Waiting',
    'Status', 'Follow-Up Sent', 'Thread ID'
  ],
  LOG_HEADERS: [
    'Timestamp', 'Action', 'To', 'Subject', 'Details'
  ],
  BRAND: {
    NAME: 'TAKScripts',
    DARK: '#1A1A1A',
    GOLD: '#C9A84C',
    WHITE: '#FFFFFF',
    LIGHT_GRAY: '#F9F9F9',
    BODY_BG: '#FAFAFA',
    SUCCESS_BG: '#E8F5E9',
    SUCCESS_TEXT: '#2E7D32',
    WARNING_BG: '#FFF8E1',
    WARNING_TEXT: '#F57F17',
    ERROR_BG: '#FFEBEE',
    ERROR_TEXT: '#C62828',
    FONT_HEADER: 'Roboto Mono',
    FONT_BODY: 'Roboto',
    FONT_SIZE: 10,
  },
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
    .addItem('\u25B6\uFE0F Start Follow-Up Nudger', 'startNudger')
    .addItem('\u23F9\uFE0F Stop Follow-Up Nudger', 'stopNudger')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run (no emails sent)', 'testRun')
    .addItem('\uD83D\uDCCA View Log', 'viewLog')
    .addSeparator()
    .addItem('\u2139\uFE0F About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  var html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('Follow-Up Nudger Settings')
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
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Follow-Up Nudger</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 \u00B7 by TAK Ventures</p>' +
      '<hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">' +
      '<p style="font-size: 13px; color: #333; line-height: 1.6;">' +
        'Never let an important email slip through the cracks.<br>' +
        'Part of the <strong>TAKScripts</strong> collection.' +
      '</p>' +
      '<p style="margin-top: 16px;">' +
        '<a href="https://takscripts.store" target="_blank" ' +
           'style="color: #C9A84C; text-decoration: none; font-weight: 600; font-size: 13px;">' +
          'takscripts.store \u2192' +
        '</a>' +
      '</p>' +
    '</div>'
  ).setWidth(300).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}

/**
 * Navigates to the follow-up log sheet.
 */
function viewLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No log entries yet. Start the Follow-Up Nudger and activity will be logged here.'
    );
    return;
  }
  ss.setActiveSheet(sheet);
}


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Returns default settings for the Follow-Up Nudger.
 */
function getDefaultSettings_() {
  return {
    daysBeforeFollowUp: 3,
    maxFollowUps: 2,
    autoSend: false,
    sendDigest: true,
    excludeDomains: 'noreply, no-reply, notifications, mailer-daemon, newsletter, calendar-notification',
    followUpSubject: 'Re: {{subject}}',
    followUpMessage:
      'Hi {{name}},\n\n' +
      'I wanted to follow up on my email regarding "{{subject}}" sent {{days}} days ago. ' +
      'I understand you may be busy, but I wanted to make sure this didn\'t slip through the cracks.\n\n' +
      'Please let me know if you have any questions or if there\'s a better time to connect.\n\n' +
      'Thanks,\n[Your Name]',
  };
}

/**
 * Save settings from the sidebar.
 * @param {Object} settings - Settings object from the sidebar form.
 * @return {Object} Success status.
 */
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('nudger_settings', JSON.stringify(settings));
  writeLog_('Settings Updated', '', '', 'Settings saved successfully');
  return { success: true };
}

/**
 * Load saved settings. Returns defaults if none are saved.
 * @return {Object} Current settings.
 */
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('nudger_settings');
  if (!raw) {
    return getDefaultSettings_();
  }
  var saved = JSON.parse(raw);
  var defaults = getDefaultSettings_();
  // Merge with defaults so new fields are always present
  for (var key in defaults) {
    if (saved[key] === undefined) {
      saved[key] = defaults[key];
    }
  }
  return saved;
}


// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main entry point — scans sent mail, identifies unreplied threads,
 * updates the tracker sheet, and optionally sends follow-ups.
 * Called by the time-driven trigger.
 */
function runFollowUpCheck() {
  try {
    var settings = loadSettings();
    var daysThreshold = parseInt(settings.daysBeforeFollowUp, 10) || 3;
    var maxFollowUps = parseInt(settings.maxFollowUps, 10) || 2;
    var autoSend = settings.autoSend === true || settings.autoSend === 'true';
    var sendDigest = settings.sendDigest === true || settings.sendDigest === 'true';

    var excludeDomains = parseExcludeDomains_(settings.excludeDomains);
    var processedIds = getProcessedThreadIds_();
    var followUpCounts = getFollowUpCounts_();

    // Search sent mail from the last N days
    var searchDays = Math.max(CONFIG.SCAN_DAYS_BACK, daysThreshold + 7);
    var query = 'in:sent newer_than:' + searchDays + 'd';
    var threads = GmailApp.search(query, 0, 200);

    var myEmail = Session.getActiveUser().getEmail().toLowerCase();
    var now = new Date();
    var pendingFollowUps = [];
    var newlySent = [];

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var threadId = thread.getId();
      var messages = thread.getMessages();

      // Find the last message I sent in this thread
      var lastSentByMe = null;
      var lastSentDate = null;
      for (var m = messages.length - 1; m >= 0; m--) {
        var msg = messages[m];
        var fromEmail = extractEmail_(msg.getFrom());
        if (fromEmail === myEmail) {
          lastSentByMe = msg;
          lastSentDate = msg.getDate();
          break;
        }
      }
      if (!lastSentByMe) continue;

      // Check if someone else replied after my last sent message
      var gotReply = false;
      for (var r = 0; r < messages.length; r++) {
        var replyMsg = messages[r];
        if (replyMsg.getDate() > lastSentDate && extractEmail_(replyMsg.getFrom()) !== myEmail) {
          gotReply = true;
          break;
        }
      }
      if (gotReply) {
        // If it was previously tracked, mark as replied
        markThreadReplied_(threadId);
        continue;
      }

      // Calculate days waiting
      var daysWaiting = Math.floor((now.getTime() - lastSentDate.getTime()) / (1000 * 60 * 60 * 24));
      if (daysWaiting < daysThreshold) continue;

      // Get recipient info
      var toField = lastSentByMe.getTo();
      var recipientEmail = extractEmail_(toField);
      var recipientName = extractName_(toField);
      var subject = lastSentByMe.getSubject();

      // Skip excluded domains
      if (shouldExclude_(recipientEmail, excludeDomains)) continue;

      // Check max follow-ups
      var currentCount = followUpCounts[threadId] || 0;
      if (currentCount >= maxFollowUps) continue;

      // Build the pending item
      var item = {
        threadId: threadId,
        dateSent: lastSentDate,
        to: toField,
        recipientEmail: recipientEmail,
        recipientName: recipientName,
        subject: subject,
        daysWaiting: daysWaiting,
        followUpsSent: currentCount,
        thread: thread,
        lastMessage: lastSentByMe,
      };

      // Auto-send follow-up if enabled and this is a new thread to nudge
      if (autoSend && !processedIds.has(threadId + '_' + (currentCount + 1))) {
        try {
          sendFollowUp_(item, settings);
          item.followUpsSent = currentCount + 1;
          incrementFollowUpCount_(threadId);
          markProcessed_(threadId, currentCount + 1);
          newlySent.push(item);
          writeLog_('Follow-Up Sent', recipientEmail, subject,
            'Auto follow-up #' + item.followUpsSent + ' after ' + daysWaiting + ' days');
        } catch (sendErr) {
          Logger.log('Error sending follow-up to ' + recipientEmail + ': ' + sendErr.message);
          writeLog_('Send Error', recipientEmail, subject, sendErr.message);
        }
      }

      pendingFollowUps.push(item);
    }

    // Update the tracker sheet
    updateTrackerSheet_(pendingFollowUps);

    // Send digest email if enabled and there are pending items
    if (sendDigest && pendingFollowUps.length > 0) {
      sendDigestEmail_(pendingFollowUps, newlySent, settings);
    }

    Logger.log('Follow-Up Nudger complete. Found ' + pendingFollowUps.length +
      ' pending, sent ' + newlySent.length + ' follow-ups.');

  } catch (err) {
    Logger.log('Follow-Up Nudger error: ' + err.message);
    writeLog_('Error', '', '', err.message);
  }
}


// ═══════════════════════════════════════════
// FOLLOW-UP EMAIL
// ═══════════════════════════════════════════

/**
 * Sends a branded follow-up email for a pending item.
 * @param {Object} item - The pending follow-up item.
 * @param {Object} settings - Current settings.
 */
function sendFollowUp_(item, settings) {
  var name = item.recipientName || 'there';
  var subject = settings.followUpSubject
    .replace(/\{\{subject\}\}/g, item.subject)
    .replace(/\{\{name\}\}/g, name)
    .replace(/\{\{days\}\}/g, String(item.daysWaiting));

  var bodyText = settings.followUpMessage
    .replace(/\{\{name\}\}/g, name)
    .replace(/\{\{subject\}\}/g, item.subject)
    .replace(/\{\{days\}\}/g, String(item.daysWaiting));

  var htmlBody = buildBrandedEmail_(bodyText, 'Follow-Up');

  // Reply to the existing thread
  item.lastMessage.reply('', {
    htmlBody: htmlBody,
    subject: subject,
  });
}

/**
 * Sends a daily digest email listing all pending follow-ups.
 * @param {Array} pending - Array of pending follow-up items.
 * @param {Array} sent - Array of items that had follow-ups sent this run.
 * @param {Object} settings - Current settings.
 */
function sendDigestEmail_(pending, sent, settings) {
  var myEmail = Session.getActiveUser().getEmail();
  var rows = '';

  for (var i = 0; i < pending.length; i++) {
    var item = pending[i];
    var statusColor = item.daysWaiting >= 7 ? '#C62828' :
                      item.daysWaiting >= 5 ? '#F57F17' : '#2E7D32';
    var wasSent = false;
    for (var s = 0; s < sent.length; s++) {
      if (sent[s].threadId === item.threadId) { wasSent = true; break; }
    }

    rows +=
      '<tr style="border-bottom: 1px solid #eee;">' +
        '<td style="padding: 10px 12px; font-size: 13px;">' + escapeHtml_(item.to) + '</td>' +
        '<td style="padding: 10px 12px; font-size: 13px;">' + escapeHtml_(item.subject) + '</td>' +
        '<td style="padding: 10px 12px; text-align: center; font-weight: 600; color: ' + statusColor + ';">' +
          item.daysWaiting + ' days</td>' +
        '<td style="padding: 10px 12px; text-align: center; font-size: 13px;">' +
          (wasSent ? '\u2705 Sent' : '\u23F3 Pending') + '</td>' +
      '</tr>';
  }

  var digestHtml =
    buildEmailHeader_('Follow-Up Digest') +
    '<div style="padding: 24px;">' +
      '<p style="font-size: 14px; color: #333; margin: 0 0 16px;">You have <strong>' +
        pending.length + '</strong> email' + (pending.length === 1 ? '' : 's') +
        ' awaiting a reply.</p>' +
      '<table style="width: 100%; border-collapse: collapse; font-family: \'Segoe UI\', sans-serif;">' +
        '<thead>' +
          '<tr style="background: #f5f5f5;">' +
            '<th style="padding: 10px 12px; text-align: left; font-size: 11px; text-transform: uppercase; color: #666;">To</th>' +
            '<th style="padding: 10px 12px; text-align: left; font-size: 11px; text-transform: uppercase; color: #666;">Subject</th>' +
            '<th style="padding: 10px 12px; text-align: center; font-size: 11px; text-transform: uppercase; color: #666;">Waiting</th>' +
            '<th style="padding: 10px 12px; text-align: center; font-size: 11px; text-transform: uppercase; color: #666;">Follow-Up</th>' +
          '</tr>' +
        '</thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +
    '</div>' +
    buildEmailFooter_();

  GmailApp.sendEmail(myEmail, '\uD83D\uDD14 Follow-Up Digest: ' + pending.length + ' pending repl' +
    (pending.length === 1 ? 'y' : 'ies'), '', {
    htmlBody: digestHtml,
    name: 'Follow-Up Nudger by TAKScripts',
  });

  writeLog_('Digest Sent', myEmail, '', pending.length + ' pending follow-ups reported');
}


// ═══════════════════════════════════════════
// BRANDED EMAIL TEMPLATES
// ═══════════════════════════════════════════

/**
 * Builds the branded email header.
 * @param {string} subtitle - The subtitle text under the brand name.
 * @return {string} HTML header string.
 */
function buildEmailHeader_(subtitle) {
  return '<!DOCTYPE html><html><body style="margin:0; padding:0; font-family: \'Segoe UI\', system-ui, sans-serif;">' +
    '<div style="max-width: 600px; margin: 0 auto; background: #fff;">' +
      '<div style="background: #1A1A1A; padding: 24px; text-align: center;">' +
        '<span style="font-size: 24px;">\uD83D\uDD77</span>' +
        '<h1 style="margin: 4px 0 0; font-size: 16px; color: #C9A84C; font-weight: 600; letter-spacing: 0.5px;">' +
          'TAKScripts</h1>' +
        '<p style="margin: 4px 0 0; font-size: 11px; color: #888;">' + escapeHtml_(subtitle) + '</p>' +
      '</div>';
}

/**
 * Builds the branded email footer.
 * @return {string} HTML footer string.
 */
function buildEmailFooter_() {
  return '<div style="background: #f5f5f5; padding: 16px; text-align: center; font-size: 11px; color: #999;">' +
      'Powered by <a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none; font-weight: 600;">' +
      'TAKScripts</a> \u00B7 takscripts.store' +
    '</div></div></body></html>';
}

/**
 * Wraps plain text body in a branded HTML email.
 * @param {string} bodyText - The plain-text message body.
 * @param {string} subtitle - Subtitle for the email header.
 * @return {string} Complete branded HTML email.
 */
function buildBrandedEmail_(bodyText, subtitle) {
  // Convert newlines to <br> for HTML
  var htmlBody = escapeHtml_(bodyText).replace(/\n/g, '<br>');
  return buildEmailHeader_(subtitle) +
    '<div style="padding: 24px; font-size: 14px; color: #333; line-height: 1.7;">' +
      htmlBody +
    '</div>' +
    buildEmailFooter_();
}


// ═══════════════════════════════════════════
// TRACKER SHEET MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Updates the follow-up tracker sheet with current pending items.
 * Creates the sheet if it doesn't exist. Clears and rewrites data each run.
 * @param {Array} items - Array of pending follow-up items.
 */
function updateTrackerSheet_(items) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    setupTrackerHeaders_(sheet);
  }

  // Clear data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length).clearContent()
      .setBackground(null).setFontColor(null);
  }

  if (items.length === 0) return;

  // Write data
  var data = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var status = item.daysWaiting >= 7 ? 'Urgent' :
                 item.daysWaiting >= 5 ? 'Overdue' : 'Pending';
    var followUpText = item.followUpsSent > 0 ?
      'Yes (' + item.followUpsSent + ')' : 'No';

    data.push([
      item.dateSent,
      item.to,
      item.subject,
      item.daysWaiting,
      status,
      followUpText,
      item.threadId,
    ]);
  }

  var dataRange = sheet.getRange(2, 1, data.length, CONFIG.HEADERS.length);
  dataRange.setValues(data);

  // Style data rows
  dataRange.setFontFamily(CONFIG.BRAND.FONT_BODY)
    .setFontSize(CONFIG.BRAND.FONT_SIZE)
    .setVerticalAlignment('middle');

  // Date format for column A
  sheet.getRange(2, 1, data.length, 1).setNumberFormat('MMM d, yyyy h:mm a');

  // Alternating row colors
  for (var row = 0; row < data.length; row++) {
    var rowRange = sheet.getRange(row + 2, 1, 1, CONFIG.HEADERS.length);
    rowRange.setBackground(row % 2 === 0 ? CONFIG.BRAND.WHITE : CONFIG.BRAND.LIGHT_GRAY);
  }

  // Status cell coloring
  for (var s = 0; s < data.length; s++) {
    var statusCell = sheet.getRange(s + 2, 5); // Column E = Status
    var statusVal = data[s][4];
    if (statusVal === 'Urgent') {
      statusCell.setBackground(CONFIG.BRAND.ERROR_BG).setFontColor(CONFIG.BRAND.ERROR_TEXT);
    } else if (statusVal === 'Overdue') {
      statusCell.setBackground(CONFIG.BRAND.WARNING_BG).setFontColor(CONFIG.BRAND.WARNING_TEXT);
    } else {
      statusCell.setBackground(CONFIG.BRAND.SUCCESS_BG).setFontColor(CONFIG.BRAND.SUCCESS_TEXT);
    }
    statusCell.setFontWeight('bold');
  }

  // Add footer row
  var footerRow = data.length + 2;
  var footerRange = sheet.getRange(footerRow, 1, 1, CONFIG.HEADERS.length);
  footerRange.merge()
    .setValue('Powered by TAKScripts \u00B7 takscripts.store')
    .setFontFamily(CONFIG.BRAND.FONT_BODY)
    .setFontSize(8)
    .setFontColor('#999')
    .setHorizontalAlignment('center')
    .setBackground(CONFIG.BRAND.WHITE);

  // Auto-resize columns
  for (var c = 1; c <= CONFIG.HEADERS.length; c++) {
    sheet.autoResizeColumn(c);
  }
  // Set minimum widths
  sheet.setColumnWidth(1, 170); // Date Sent
  sheet.setColumnWidth(2, 220); // To
  sheet.setColumnWidth(3, 280); // Subject
  sheet.setColumnWidth(4, 100); // Days Waiting
  sheet.setColumnWidth(5, 100); // Status
  sheet.setColumnWidth(6, 120); // Follow-Up Sent
  sheet.setColumnWidth(7, 160); // Thread ID
}

/**
 * Sets up the header row for the tracker sheet.
 * @param {Sheet} sheet - The Google Sheet to set up.
 */
function setupTrackerHeaders_(sheet) {
  sheet.appendRow(CONFIG.HEADERS);
  var headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground(CONFIG.BRAND.DARK)
    .setFontColor(CONFIG.BRAND.GOLD)
    .setFontFamily(CONFIG.BRAND.FONT_HEADER)
    .setFontSize(CONFIG.BRAND.FONT_SIZE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);
}


// ═══════════════════════════════════════════
// ACTIVITY LOG
// ═══════════════════════════════════════════

/**
 * Writes an entry to the activity log sheet.
 * @param {string} action - The action taken.
 * @param {string} to - Recipient email (if applicable).
 * @param {string} subject - Email subject (if applicable).
 * @param {string} details - Additional details.
 */
function writeLog_(action, to, subject, details) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      sheet.appendRow(CONFIG.LOG_HEADERS);
      var headerRange = sheet.getRange(1, 1, 1, CONFIG.LOG_HEADERS.length);
      headerRange
        .setFontWeight('bold')
        .setBackground(CONFIG.BRAND.DARK)
        .setFontColor(CONFIG.BRAND.GOLD)
        .setFontFamily(CONFIG.BRAND.FONT_HEADER)
        .setFontSize(CONFIG.BRAND.FONT_SIZE);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 140);
      sheet.setColumnWidth(3, 220);
      sheet.setColumnWidth(4, 280);
      sheet.setColumnWidth(5, 300);
    }

    sheet.appendRow([new Date(), action, to, subject, details]);
  } catch (logErr) {
    Logger.log('Log write error: ' + logErr.message);
  }
}


// ═══════════════════════════════════════════
// TRACKING STATE (IDEMPOTENCY)
// ═══════════════════════════════════════════

/**
 * Gets the set of already-processed thread IDs (with follow-up count suffix).
 * @return {Object} A Set-like object of processed keys.
 */
function getProcessedThreadIds_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('nudger_processed') || '[]';
  var arr = JSON.parse(raw);
  var set = {};
  for (var i = 0; i < arr.length; i++) {
    set[arr[i]] = true;
  }
  // Return a Set-like wrapper
  return {
    _data: set,
    has: function(key) { return this._data[key] === true; },
    add: function(key) { this._data[key] = true; },
    toArray: function() { return Object.keys(this._data); },
  };
}

/**
 * Marks a thread + follow-up count as processed.
 * @param {string} threadId - The Gmail thread ID.
 * @param {number} count - The follow-up number.
 */
function markProcessed_(threadId, count) {
  var processed = getProcessedThreadIds_();
  processed.add(threadId + '_' + count);
  var props = PropertiesService.getScriptProperties();
  props.setProperty('nudger_processed', JSON.stringify(processed.toArray()));
}

/**
 * Gets follow-up counts per thread ID.
 * @return {Object} Map of threadId to follow-up count.
 */
function getFollowUpCounts_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('nudger_followup_counts') || '{}';
  return JSON.parse(raw);
}

/**
 * Increments the follow-up count for a thread.
 * @param {string} threadId - The Gmail thread ID.
 */
function incrementFollowUpCount_(threadId) {
  var counts = getFollowUpCounts_();
  counts[threadId] = (counts[threadId] || 0) + 1;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('nudger_followup_counts', JSON.stringify(counts));
}

/**
 * Marks a thread as replied (removes from tracking).
 * @param {string} threadId - The Gmail thread ID.
 */
function markThreadReplied_(threadId) {
  // Clean up follow-up counts for this thread
  var counts = getFollowUpCounts_();
  if (counts[threadId]) {
    delete counts[threadId];
    var props = PropertiesService.getScriptProperties();
    props.setProperty('nudger_followup_counts', JSON.stringify(counts));
  }
}


// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

/**
 * Extracts the email address from a "Name <email>" formatted string.
 * @param {string} fromField - The From or To header value.
 * @return {string} Lowercase email address.
 */
function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

/**
 * Extracts the first name from a "Name <email>" formatted string.
 * @param {string} fromField - The From or To header value.
 * @return {string} First name, or empty string.
 */
function extractName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim().split(' ')[0] : '';
}

/**
 * Parses the exclude domains setting into an array.
 * @param {string} domainsStr - Comma-separated domain/keyword list.
 * @return {Array<string>} Array of lowercase domain keywords.
 */
function parseExcludeDomains_(domainsStr) {
  return (domainsStr || '')
    .split(',')
    .map(function(d) { return d.trim().toLowerCase(); })
    .filter(function(d) { return d.length > 0; });
}

/**
 * Checks if an email should be excluded based on domain rules.
 * @param {string} email - The email address to check.
 * @param {Array<string>} excludeDomains - Domains/keywords to exclude.
 * @return {boolean} True if the email should be excluded.
 */
function shouldExclude_(email, excludeDomains) {
  for (var i = 0; i < excludeDomains.length; i++) {
    if (email.indexOf(excludeDomains[i]) !== -1) return true;
  }
  return false;
}

/**
 * Escapes HTML special characters.
 * @param {string} str - Raw string.
 * @return {string} HTML-safe string.
 */
function escapeHtml_(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


// ═══════════════════════════════════════════
// CONTROLS
// ═══════════════════════════════════════════

/**
 * Starts the Follow-Up Nudger. Creates an hourly trigger.
 */
function startNudger() {
  var settings = loadSettings();

  // Remove any existing triggers first
  stopNudger_(true);

  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION)
    .timeBased()
    .everyHours(CONFIG.TRIGGER_INTERVAL_HOURS)
    .create();

  // Run immediately
  runFollowUpCheck();

  writeLog_('Nudger Started', '', '', 'Trigger created, running every ' +
    CONFIG.TRIGGER_INTERVAL_HOURS + ' hour(s)');

  SpreadsheetApp.getUi().alert(
    '\u2705 Follow-Up Nudger is ACTIVE\n\n' +
    'Checking every ' + CONFIG.TRIGGER_INTERVAL_HOURS + ' hour(s) for unreplied emails.\n' +
    'Follow-ups after: ' + settings.daysBeforeFollowUp + ' days\n' +
    'Auto-send: ' + (settings.autoSend === true || settings.autoSend === 'true' ? 'ON' : 'OFF') + '\n' +
    'Email digest: ' + (settings.sendDigest === true || settings.sendDigest === 'true' ? 'ON' : 'OFF') + '\n\n' +
    'To stop: \uD83D\uDD77 TAKScripts \u2192 Stop Follow-Up Nudger'
  );
}

/**
 * Stops the Follow-Up Nudger and cleans up triggers.
 * @param {boolean} [silent] - If true, don't show UI alert.
 */
function stopNudger_(silent) {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === CONFIG.TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  if (!silent) {
    writeLog_('Nudger Stopped', '', '', removed + ' trigger(s) removed');
    SpreadsheetApp.getUi().alert(
      '\u23F9 Follow-Up Nudger has been stopped.\n\n' +
      'No more follow-ups will be sent automatically.\n' +
      'Your tracker sheet data has been preserved.'
    );
  }
}

/**
 * Public stop function (shows UI alert).
 */
function stopNudger() {
  stopNudger_(false);
}

/**
 * Dry run — shows what would be flagged without sending any emails.
 */
function testRun() {
  try {
    var settings = loadSettings();
    var daysThreshold = parseInt(settings.daysBeforeFollowUp, 10) || 3;
    var excludeDomains = parseExcludeDomains_(settings.excludeDomains);
    var maxFollowUps = parseInt(settings.maxFollowUps, 10) || 2;
    var followUpCounts = getFollowUpCounts_();

    var searchDays = Math.max(CONFIG.SCAN_DAYS_BACK, daysThreshold + 7);
    var query = 'in:sent newer_than:' + searchDays + 'd';
    var threads = GmailApp.search(query, 0, 200);

    var myEmail = Session.getActiveUser().getEmail().toLowerCase();
    var now = new Date();
    var results = [];

    results.push('\uD83E\uDDEA TEST RUN \u2014 No emails will be sent');
    results.push('Settings: follow up after ' + daysThreshold + ' days, max ' + maxFollowUps + ' follow-ups');
    results.push('Scanned ' + threads.length + ' sent threads from last ' + searchDays + ' days');
    results.push('---');

    var pendingCount = 0;

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var messages = thread.getMessages();

      var lastSentByMe = null;
      var lastSentDate = null;
      for (var m = messages.length - 1; m >= 0; m--) {
        var msg = messages[m];
        if (extractEmail_(msg.getFrom()) === myEmail) {
          lastSentByMe = msg;
          lastSentDate = msg.getDate();
          break;
        }
      }
      if (!lastSentByMe) continue;

      var gotReply = false;
      for (var r = 0; r < messages.length; r++) {
        if (messages[r].getDate() > lastSentDate && extractEmail_(messages[r].getFrom()) !== myEmail) {
          gotReply = true;
          break;
        }
      }
      if (gotReply) continue;

      var daysWaiting = Math.floor((now.getTime() - lastSentDate.getTime()) / (1000 * 60 * 60 * 24));
      if (daysWaiting < daysThreshold) continue;

      var toField = lastSentByMe.getTo();
      var recipientEmail = extractEmail_(toField);
      var subject = lastSentByMe.getSubject();

      var excluded = shouldExclude_(recipientEmail, excludeDomains);
      var currentCount = followUpCounts[thread.getId()] || 0;
      var maxedOut = currentCount >= maxFollowUps;

      results.push('To: ' + toField);
      results.push('Subject: ' + subject);
      results.push('Days waiting: ' + daysWaiting);

      if (excluded) {
        results.push('\u2192 SKIP (domain excluded)');
      } else if (maxedOut) {
        results.push('\u2192 SKIP (max follow-ups reached: ' + currentCount + '/' + maxFollowUps + ')');
      } else {
        results.push('\u2192 WOULD FLAG for follow-up' +
          (settings.autoSend === true || settings.autoSend === 'true' ? ' and AUTO-SEND' : ''));
        pendingCount++;
      }
      results.push('');
    }

    results.push('---');
    results.push('Total pending follow-ups: ' + pendingCount);

    if (pendingCount === 0) {
      results.push('All clear! No emails need follow-up right now.');
    }

    var output = results.join('\n');
    Logger.log(output);
    SpreadsheetApp.getUi().alert(output);

    writeLog_('Test Run', '', '', pendingCount + ' pending follow-ups found');

  } catch (err) {
    SpreadsheetApp.getUi().alert('Error during test run: ' + err.message);
    Logger.log('Test run error: ' + err.message);
  }
}


// ═══════════════════════════════════════════
// SETTINGS SIDEBAR HTML
// ═══════════════════════════════════════════

/**
 * Returns the complete HTML for the settings sidebar.
 * @return {string} HTML string for the sidebar.
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
'      color: #1A1A1A;' +
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
'    .field input, .field textarea, .field select {' +
'      width: 100%;' +
'      padding: 10px 12px;' +
'      border: 1px solid #ddd;' +
'      border-radius: 6px;' +
'      font-size: 13px;' +
'      font-family: inherit;' +
'      transition: border-color 0.2s, box-shadow 0.2s;' +
'      background: #fff;' +
'    }' +
'    .field input:focus, .field textarea:focus, .field select:focus {' +
'      outline: none;' +
'      border-color: #C9A84C;' +
'      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);' +
'    }' +
'    .field textarea { min-height: 120px; resize: vertical; line-height: 1.5; }' +
'    .field .help {' +
'      font-size: 11px;' +
'      color: #999;' +
'      margin-top: 4px;' +
'      line-height: 1.4;' +
'    }' +
'    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }' +
'    .toggle-wrap {' +
'      display: flex;' +
'      align-items: center;' +
'      justify-content: space-between;' +
'      padding: 12px;' +
'      border: 1px solid #eee;' +
'      border-radius: 8px;' +
'      margin-bottom: 10px;' +
'      background: #fff;' +
'    }' +
'    .toggle-wrap .label {' +
'      font-size: 13px;' +
'      font-weight: 500;' +
'    }' +
'    .toggle-wrap .sublabel {' +
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
'      height: 18px; width: 18px;' +
'      left: 3px; bottom: 3px;' +
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
'    .variables {' +
'      background: #f5f5f5;' +
'      border-radius: 6px;' +
'      padding: 10px 12px;' +
'      margin-top: 8px;' +
'    }' +
'    .variables code {' +
'      display: inline-block;' +
'      background: #e8e8e8;' +
'      padding: 1px 6px;' +
'      border-radius: 3px;' +
'      font-size: 11px;' +
'      margin: 2px;' +
'    }' +
'    .section-title {' +
'      font-size: 11px;' +
'      font-weight: 700;' +
'      text-transform: uppercase;' +
'      letter-spacing: 1.5px;' +
'      color: #C9A84C;' +
'      margin: 20px 0 12px;' +
'      padding-bottom: 6px;' +
'      border-bottom: 1px solid #eee;' +
'    }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <div class="logo">\uD83D\uDD77</div>' +
'    <h1><span class="brand">TAK</span>Scripts</h1>' +
'    <div class="sub">Follow-Up Nudger \u00B7 Settings</div>' +
'  </div>' +
'' +
'  <div class="form">' +
'    <div class="section-title">Timing</div>' +
'' +
'    <div class="row">' +
'      <div class="field">' +
'        <label>Days Before Follow-Up</label>' +
'        <input type="number" id="daysBeforeFollowUp" min="1" max="30" value="3" />' +
'        <div class="help">Wait this many days</div>' +
'      </div>' +
'      <div class="field">' +
'        <label>Max Follow-Ups</label>' +
'        <input type="number" id="maxFollowUps" min="1" max="10" value="2" />' +
'        <div class="help">Per thread</div>' +
'      </div>' +
'    </div>' +
'' +
'    <div class="section-title">Behavior</div>' +
'' +
'    <div class="toggle-wrap">' +
'      <div>' +
'        <div class="label">Auto-Send Follow-Ups</div>' +
'        <div class="sublabel">Automatically send nudge emails</div>' +
'      </div>' +
'      <label class="switch">' +
'        <input type="checkbox" id="autoSend" />' +
'        <span class="slider"></span>' +
'      </label>' +
'    </div>' +
'' +
'    <div class="toggle-wrap">' +
'      <div>' +
'        <div class="label">Email Digest</div>' +
'        <div class="sublabel">Get a summary of pending follow-ups</div>' +
'      </div>' +
'      <label class="switch">' +
'        <input type="checkbox" id="sendDigest" checked />' +
'        <span class="slider"></span>' +
'      </label>' +
'    </div>' +
'' +
'    <div class="section-title">Follow-Up Message</div>' +
'' +
'    <div class="field">' +
'      <label>Subject Line</label>' +
'      <input type="text" id="followUpSubject" placeholder="Re: {{subject}}" />' +
'    </div>' +
'' +
'    <div class="field">' +
'      <label>Message Template</label>' +
'      <textarea id="followUpMessage" placeholder="Hi {{name}}, I wanted to follow up..."></textarea>' +
'      <div class="variables">' +
'        <div class="help" style="margin-bottom: 4px;">Template variables:</div>' +
'        <code>{{name}}</code> recipient\'s first name<br>' +
'        <code>{{subject}}</code> original subject<br>' +
'        <code>{{days}}</code> days since you sent it' +
'      </div>' +
'    </div>' +
'' +
'    <div class="section-title">Filters</div>' +
'' +
'    <div class="field">' +
'      <label>Exclude Domains / Keywords</label>' +
'      <input type="text" id="excludeDomains" placeholder="noreply, newsletter, mailer-daemon" />' +
'      <div class="help">Comma-separated. Emails containing these words will be skipped.</div>' +
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
'      document.getElementById("daysBeforeFollowUp").value = settings.daysBeforeFollowUp || 3;' +
'      document.getElementById("maxFollowUps").value = settings.maxFollowUps || 2;' +
'      document.getElementById("autoSend").checked = settings.autoSend === true || settings.autoSend === "true";' +
'      document.getElementById("sendDigest").checked = settings.sendDigest !== false && settings.sendDigest !== "false";' +
'      document.getElementById("followUpSubject").value = settings.followUpSubject || "";' +
'      document.getElementById("followUpMessage").value = settings.followUpMessage || "";' +
'      document.getElementById("excludeDomains").value = settings.excludeDomains || "";' +
'    }).loadSettings();' +
'' +
'    function save() {' +
'      var settings = {' +
'        daysBeforeFollowUp: parseInt(document.getElementById("daysBeforeFollowUp").value, 10) || 3,' +
'        maxFollowUps: parseInt(document.getElementById("maxFollowUps").value, 10) || 2,' +
'        autoSend: document.getElementById("autoSend").checked,' +
'        sendDigest: document.getElementById("sendDigest").checked,' +
'        followUpSubject: document.getElementById("followUpSubject").value,' +
'        followUpMessage: document.getElementById("followUpMessage").value,' +
'        excludeDomains: document.getElementById("excludeDomains").value,' +
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
