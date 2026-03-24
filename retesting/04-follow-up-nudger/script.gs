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
 * - Tracks pending follow-ups in a branded "Follow-Up Dashboard" sheet
 * - Dashboard stats: Pending · Overdue · Avg Wait · Auto-Sent · Longest
 * - Custom follow-up message template with {{name}}, {{subject}}, {{days}} variables
 * - Exclude specific domains from tracking
 * - Configurable max follow-ups per thread
 * - Auto-send follow-up emails (optional) or email digest notification
 * - Test Run mode — preview what would be flagged without sending anything
 * - Runs on schedule via time-driven trigger (every hour)
 * - Idempotent — tracks thread IDs already processed
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

var DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Follow-Up Dashboard';
var LOG_SHEET_NAME = '\uD83D\uDCDD Follow-Up Log';
var DASHBOARD_HEADER_ROW = 4;

var BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E6F4EA', successText: '#137333',
  warningBg: '#FEF7E0', warningText: '#B06000',
  errorBg: '#FCE8E6', errorText: '#C5221F',
  infoBg: '#E3F2FD', infoText: '#1565C0',
  headerFont: 'Roboto Mono', bodyFont: 'Roboto',
};

var COL = {
  dateSent: 1,
  to: 2,
  subject: 3,
  daysWaiting: 4,
  status: 5,
  followUpSent: 6,
  threadId: 7,
  notes: 8,
  snoozeUntil: 9,
  gmailLink: 10,
};

var HEADERS = [
  'Date Sent', 'To', 'Subject', 'Days Waiting',
  'Status', 'Follow-Up Sent', 'Thread ID',
  'Notes', 'Snooze Until', 'Gmail Link',
];

var TRIGGER_FUNCTION = 'runFollowUpCheck';
var TRIGGER_INTERVAL_HOURS = 1;
var SCAN_DAYS_BACK = 14;


// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

/**
 * Creates the TAKScripts menu when the spreadsheet opens.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('\uD83D\uDD77 TAKScripts')
    .addItem('\u2699\uFE0F Settings', 'showSettings')
    .addItem('\u25B6\uFE0F Start Follow-Up Nudger', 'startNudger')
    .addItem('\u23F9\uFE0F Stop Follow-Up Nudger', 'stopNudger')
    .addSeparator()
    .addItem('\uD83E\uDDEA Test Run (no emails sent)', 'testRun')
    .addItem('\uD83D\uDCCA View Dashboard', 'viewDashboard')
    .addItem('\u2713 Mark as Resolved', 'markResolved')
    .addItem('\uD83D\uDCA4 Snooze Selected (7 days)', 'snoozeRow')
    .addItem('\uD83D\uDD04 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('\u2753 How to Use', 'showHelp')
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
 * Navigates to the Follow-Up Dashboard sheet.
 */
function viewDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = getOrCreateDashboard_();
  }
  ss.setActiveSheet(sheet);
}

/**
 * Shows the about dialog.
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
      '<div style="font-size: 32px; margin-bottom: 8px;">\uD83D\uDD77</div>' +
      '<h2 style="margin: 0 0 4px; font-size: 18px;">Follow-Up Nudger</h2>' +
      '<p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 \u00B7 by TAK Ventures</p>' +
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
  try {
    SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
  } catch(e) {
    Logger.log('About TAKScripts');
  }
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
    overdueAfterDays: 5,
    urgentAfterDays: 7,
    autoSend: false,
    sendDigest: true,
    autoSchedule: false,
    triggerIntervalHours: 1,
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
  for (var key in defaults) {
    if (saved[key] === undefined) {
      saved[key] = defaults[key];
    }
  }
  return saved;
}


// ═══════════════════════════════════════════
// DASHBOARD MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Gets or creates the Follow-Up Dashboard sheet with stats bar and headers.
 * Idempotent — safe to call multiple times.
 * @return {Sheet} The dashboard sheet.
 */
function getOrCreateDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    var oldSheet = ss.getSheetByName('\uD83D\uDCEC Follow-Up Tracker');
    if (oldSheet) {
      oldSheet.setName(DASHBOARD_SHEET_NAME);
      oldSheet.clear();
      sheet = oldSheet;
    } else {
      sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
    }
  } else {
    var headerCheck = sheet.getRange(DASHBOARD_HEADER_ROW, 1).getValue();
    if (headerCheck === HEADERS[0]) {
      return sheet; // Already set up
    }
  }

  var numCols = HEADERS.length; // 10

  // ── Row 1: Title bar ────────────────────────────────────
  sheet.getRange(1, 1, 1, numCols)
    .merge()
    .setValue('\uD83D\uDCCA  FOLLOW-UP DASHBOARD')
    .setFontFamily(BRAND.headerFont)
    .setFontSize(13)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground('#0D0D0D')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 44);

  // ── Row 2: Stat values ──────────────────────────────────
  sheet.getRange(2, 1, 1, numCols)
    .setValues([['—', '—', '—', '—', '—', '', '', '', '', '']])
    .setFontFamily(BRAND.headerFont)
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(BRAND.gold)
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 60);

  // Gold accent bottom border on stat value row
  sheet.getRange(2, 1, 1, numCols)
    .setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ── Row 3: Stat labels ──────────────────────────────────
  sheet.getRange(3, 1, 1, numCols)
    .setValues([['PENDING', 'OVERDUE', 'AVG WAIT', 'AUTO-SENT', 'LONGEST', 'LAST RUN', '', '', '', '']])
    .setFontFamily(BRAND.headerFont)
    .setFontSize(8)
    .setFontWeight('normal')
    .setFontColor('#888888')
    .setBackground(BRAND.darkBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 22);

  // ── Row 4: Column headers ───────────────────────────────
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

  // Freeze rows 1–4 (title bar + stat bar + col headers)
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // ── Column widths ───────────────────────────────────────
  sheet.setColumnWidth(COL.dateSent, 170);
  sheet.setColumnWidth(COL.to, 220);
  sheet.setColumnWidth(COL.subject, 280);
  sheet.setColumnWidth(COL.daysWaiting, 110);
  sheet.setColumnWidth(COL.status, 100);
  sheet.setColumnWidth(COL.followUpSent, 130);
  sheet.setColumnWidth(COL.threadId, 160);
  sheet.setColumnWidth(COL.notes, 200);
  sheet.setColumnWidth(COL.snoozeUntil, 130);
  sheet.setColumnWidth(COL.gmailLink, 100);
  sheet.hideColumns(COL.threadId); // Internal dedup field

  // Hide gridlines
  ss.setHiddenGridlines(true);

  return sheet;
}

/**
 * Recalculates all stats from current data rows and updates the stats bar.
 * Safe to call any time — reads live data.
 */
function refreshDashboardStats() {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;
  var lastRow = sheet.getLastRow();

  var pending = 0;
  var overdue = 0;
  var totalWait = 0;
  var autoSentCount = 0;
  var longest = 0;

  if (lastRow >= dataStartRow) {
    var numRows = lastRow - dataStartRow + 1;
    var data = sheet.getRange(dataStartRow, 1, numRows, HEADERS.length).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[COL.dateSent - 1]) continue; // Skip empty rows

      var daysWaiting = parseInt(row[COL.daysWaiting - 1], 10) || 0;
      var status = row[COL.status - 1];
      var followUpSent = row[COL.followUpSent - 1];

      // Handle snooze
      var snoozeVal = row[COL.snoozeUntil - 1];
      var actualRow = dataStartRow + i;
      if (snoozeVal) {
        var snoozeDate = new Date(snoozeVal);
        if (snoozeDate > new Date()) {
          continue; // Still snoozed — skip from stats
        } else {
          // Snooze expired — clear it and un-hide
          sheet.getRange(actualRow, COL.snoozeUntil).clearContent();
          sheet.showRows(actualRow);
        }
      }

      pending++;
      totalWait += daysWaiting;
      if (status === 'Overdue' || status === 'Urgent') overdue++;
      if (followUpSent && followUpSent.toString().indexOf('Yes') === 0) autoSentCount++;
      if (daysWaiting > longest) longest = daysWaiting;
    }
  }

  var avgWait = pending > 0 ? Math.round(totalWait / pending) : 0;
  var lastRunStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a');

  // Write stat values to row 2
  sheet.getRange(2, 1, 1, 6).setValues([[
    pending || '—',
    overdue || '—',
    pending > 0 ? avgWait + 'd' : '—',
    autoSentCount || '—',
    longest > 0 ? longest + 'd' : '—',
    lastRunStr,
  ]]);

  // Ensure LAST RUN label stays in row 3
  sheet.getRange(3, 6).setValue('LAST RUN');

  // Hide resolved rows
  for (var r = dataStartRow; r <= sheet.getLastRow(); r++) {
    var statusVal = sheet.getRange(r, COL.status).getValue().toString();
    if (statusVal.indexOf('Done') !== -1 || statusVal === '\u2713') {
      sheet.hideRows(r);
    }
  }
}


/**
 * Sets the selected row's status to "Done ✓" and hides it on next refresh.
 */
function markResolved() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;
  var row = ss.getActiveRange().getRow();
  if (row <= DASHBOARD_HEADER_ROW) {
    try { SpreadsheetApp.getUi().alert('Select a data row in the dashboard first.'); } catch(e) { Logger.log('Select a data row in the dashboard first.'); }
    return;
  }
  sheet.getRange(row, COL.status).setValue('Done \u2713');
  sheet.getRange(row, COL.status).setFontColor('#888888');
  refreshDashboardStats();
}


/**
 * Snoozes the selected dashboard row for 7 days.
 * Sets the Snooze Until date so the row is hidden during that period.
 */
function snoozeRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;
  var row = ss.getActiveRange().getRow();
  if (row <= DASHBOARD_HEADER_ROW) {
    try { SpreadsheetApp.getUi().alert('Select a data row in the dashboard first.'); }
    catch(e) { Logger.log('Select a data row in the dashboard first.'); }
    return;
  }
  var snoozeDate = new Date();
  snoozeDate.setDate(snoozeDate.getDate() + 7);
  sheet.getRange(row, COL.snoozeUntil).setValue(snoozeDate);
  sheet.getRange(row, COL.snoozeUntil).setNumberFormat('yyyy-mm-dd');
  sheet.hideRows(row);
  refreshDashboardStats();
  try {
    SpreadsheetApp.getUi().alert('\uD83D\uDCA4 Snoozed until ' +
      Utilities.formatDate(snoozeDate, Session.getScriptTimeZone(), 'MMM d') +
      '. Row hidden. It will reappear on next refresh after the snooze expires.');
  } catch(e) {
    Logger.log('Snoozed row ' + row);
  }
}


// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main entry point — scans sent mail, identifies unreplied threads,
 * updates the dashboard sheet, and optionally sends follow-ups.
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

    var searchDays = Math.max(SCAN_DAYS_BACK, daysThreshold + 7);
    var query = 'in:sent newer_than:' + searchDays + 'd';
    var threads = GmailApp.search(query, 0, 200);

    var myEmail = Session.getActiveUser().getEmail().toLowerCase();
    var now = new Date();
    var pendingFollowUps = [];
    var newlySent = [];

    for (var t = 0; t < threads.length; t++) {
      try {
        var thread = threads[t];
        var threadId = thread.getId();
        var allMessages = thread.getMessages();

        var lastSentByMe = null;
        var lastSentDate = null;
        for (var m = allMessages.length - 1; m >= 0; m--) {
          var msg = allMessages[m];
          var fromEmail = extractEmail_(msg.getFrom());
          if (fromEmail === myEmail) {
            lastSentByMe = msg;
            lastSentDate = msg.getDate();
            break;
          }
        }
        if (!lastSentByMe) continue;

        // Only look at messages after the sent date to avoid scanning full thread history
        var messages = allMessages.filter(function(m) {
          return m.getDate() >= lastSentDate;
        });
        if (messages.length === 0) messages = allMessages;

        var gotReply = false;
        for (var r = 0; r < messages.length; r++) {
          var replyMsg = messages[r];
          if (replyMsg.getDate() > lastSentDate && extractEmail_(replyMsg.getFrom()) !== myEmail) {
            gotReply = true;
            break;
          }
        }
        if (gotReply) {
          markThreadReplied_(threadId);
          continue;
        }

        var daysWaiting = Math.floor((now.getTime() - lastSentDate.getTime()) / (1000 * 60 * 60 * 24));
        if (daysWaiting < daysThreshold) continue;

        var toField = lastSentByMe.getTo();
        var recipientEmail = extractEmail_(toField);
        var recipientName = extractName_(toField);
        var subject = lastSentByMe.getSubject();

        if (shouldExclude_(recipientEmail, excludeDomains)) continue;

        var currentCount = followUpCounts[threadId] || 0;
        if (currentCount >= maxFollowUps) continue;

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
      } catch (e) {
        Logger.log('Thread error: ' + e.message);
      }
      Utilities.sleep(100);
    }

    updateTrackerSheet_(pendingFollowUps);

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
  item.lastMessage.reply('', {
    htmlBody: htmlBody,
    subject: subject,
  });
}

/**
 * Sends a daily digest email listing all pending follow-ups.
 */
function sendDigestEmail_(pending, sent, settings) {
  var myEmail = Session.getActiveUser().getEmail();
  var rows = '';
  for (var i = 0; i < pending.length; i++) {
    var item = pending[i];
    var wasSent = false;
    for (var s = 0; s < sent.length; s++) {
      if (sent[s].threadId === item.threadId) { wasSent = true; break; }
    }
    var statusColor = item.daysWaiting >= settings.urgentAfterDays ? BRAND.errorText :
      item.daysWaiting >= settings.overdueAfterDays ? BRAND.warningText : BRAND.infoText;
    var statusLabel = item.daysWaiting >= settings.urgentAfterDays ? 'Urgent' :
      item.daysWaiting >= settings.overdueAfterDays ? 'Overdue' : 'Pending';
    var statusBadge = '<span style="background:' + statusColor + ';color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">' + statusLabel + '</span>';
    rows +=
      '<tr style="border-bottom: 1px solid #eee;">' +
      '<td style="padding: 10px 12px; font-size: 13px;">' + escapeHtml_(item.to) + '</td>' +
      '<td style="padding: 10px 12px; font-size: 13px;">' + escapeHtml_(item.subject) + '</td>' +
      '<td style="padding: 10px 12px; text-align: center; font-weight: 600; color: ' + statusColor + ';">' +
        item.daysWaiting + ' days</td>' +
      '<td style="padding: 10px 12px; text-align: center;">' + statusBadge + '</td>' +
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
    '<th style="padding: 10px 12px; text-align: center; font-size: 11px; text-transform: uppercase; color: #666;">Status</th>' +
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

function buildEmailFooter_() {
  return '<div style="background: #f5f5f5; padding: 16px; text-align: center; font-size: 11px; color: #999;">' +
    'Powered by <a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none; font-weight: 600;">' +
    'TAKScripts</a> \u00B7 takscripts.store' +
    '</div></div></body></html>';
}

function buildBrandedEmail_(bodyText, subtitle) {
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
 * Updates the dashboard sheet with current pending items.
 * Clears data rows and rewrites fresh each run.
 * @param {Array} items - Array of pending follow-up items.
 */
function updateTrackerSheet_(items) {
  var sheet = getOrCreateDashboard_();
  var dataStartRow = DASHBOARD_HEADER_ROW + 1;

  // Clear existing data rows (preserve title bar + stats + headers)
  var lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, HEADERS.length)
      .clearContent()
      .setBackground(null)
      .setFontColor(null)
      .setFontWeight('normal');
  }

  if (items.length === 0) {
    refreshDashboardStats();
    return;
  }

  // Load settings once — passed to styleDataRow_ to avoid per-row PropertiesService reads
  var settings = loadSettings();

  // Write data rows
  var data = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var status = item.daysWaiting >= settings.urgentAfterDays ? 'Urgent' :
      item.daysWaiting >= settings.overdueAfterDays ? 'Overdue' : 'Pending';
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
      '',
      '',
      'https://mail.google.com/mail/u/0/#all/' + item.threadId,
    ]);
  }

  var dataRange = sheet.getRange(dataStartRow, 1, data.length, HEADERS.length);
  dataRange.setValues(data);

  // Base formatting
  dataRange
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setFontColor('#333333');
  sheet.getRange(dataStartRow, COL.dateSent, data.length, 1)
    .setNumberFormat('MMM d, yyyy');

  // Style each row (settings passed in — no per-row PropertiesService reads)
  for (var row = 0; row < data.length; row++) {
    styleDataRow_(sheet, dataStartRow + row, data[row], settings);
  }

  // Auto-resize key columns with minimum widths enforced
  sheet.autoResizeColumn(COL.to);
  sheet.autoResizeColumn(COL.subject);
  if (sheet.getColumnWidth(COL.to) < 220) sheet.setColumnWidth(COL.to, 220);
  if (sheet.getColumnWidth(COL.subject) < 280) sheet.setColumnWidth(COL.subject, 280);

  refreshDashboardStats();
}

/**
 * Applies alternating row background and status/days coloring to a data row.
 * @param {Sheet} sheet - The dashboard sheet.
 * @param {number} rowNum - 1-based row number.
 * @param {Array} rowData - Row data array.
 * @param {Object} settings - Pre-loaded settings object (avoids per-row PropertiesService reads).
 */
function styleDataRow_(sheet, rowNum, rowData, settings) {
  var isEven = (rowNum - DASHBOARD_HEADER_ROW) % 2 === 0;
  var baseBg = isEven ? BRAND.lightGray : BRAND.white;
  sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground(baseBg);
  sheet.setRowHeight(rowNum, 30);

  // Status cell (col 5)
  var status = rowData[COL.status - 1];
  var statusCell = sheet.getRange(rowNum, COL.status);
  if (status === 'Urgent') {
    statusCell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText);
  } else if (status === 'Overdue') {
    statusCell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText);
  } else {
    statusCell.setBackground(BRAND.successBg).setFontColor(BRAND.successText);
  }
  statusCell.setFontWeight('bold').setHorizontalAlignment('center');

  // Days Waiting cell (col 4) — color-coded by urgency
  var daysWaiting = parseInt(rowData[COL.daysWaiting - 1], 10) || 0;
  var daysCell = sheet.getRange(rowNum, COL.daysWaiting);
  if (daysWaiting >= settings.urgentAfterDays) {
    daysCell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText).setFontWeight('bold');
  } else if (daysWaiting >= settings.overdueAfterDays) {
    daysCell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText).setFontWeight('bold');
  }
  daysCell.setHorizontalAlignment('center');
}


// ═══════════════════════════════════════════
// ACTIVITY LOG
// ═══════════════════════════════════════════

var LOG_HEADERS_LIST = ['Timestamp', 'Action', 'To', 'Subject', 'Details'];

/**
 * Writes an entry to the activity log sheet.
 */
function writeLog_(action, to, subject, details) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(LOG_HEADERS_LIST);
      var headerRange = sheet.getRange(1, 1, 1, LOG_HEADERS_LIST.length);
      headerRange
        .setFontWeight('bold')
        .setBackground(BRAND.darkBg)
        .setFontColor(BRAND.gold)
        .setFontFamily(BRAND.headerFont)
        .setFontSize(9);
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

function getProcessedThreadIds_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('nudger_processed') || '[]';
  var arr = JSON.parse(raw);
  var set = {};
  for (var i = 0; i < arr.length; i++) {
    set[arr[i]] = true;
  }
  return {
    _data: set,
    has: function(key) { return this._data[key] === true; },
    add: function(key) { this._data[key] = true; },
    toArray: function() { return Object.keys(this._data); },
  };
}

function markProcessed_(threadId, count) {
  var processed = getProcessedThreadIds_();
  processed.add(threadId + '_' + count);
  var props = PropertiesService.getScriptProperties();
  var arr = processed.toArray();
  var MAX_IDS = 3000;
  if (arr.length > MAX_IDS) {
    arr = arr.slice(arr.length - MAX_IDS);
  }
  props.setProperty('nudger_processed', JSON.stringify(arr));
}

function getFollowUpCounts_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('nudger_followup_counts') || '{}';
  return JSON.parse(raw);
}

function incrementFollowUpCount_(threadId) {
  var counts = getFollowUpCounts_();
  counts[threadId] = (counts[threadId] || 0) + 1;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('nudger_followup_counts', JSON.stringify(counts));
}

function markThreadReplied_(threadId) {
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

function extractEmail_(fromField) {
  var match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

function extractName_(fromField) {
  var match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim().split(' ')[0] : '';
}

function parseExcludeDomains_(domainsStr) {
  return (domainsStr || '')
    .split(',')
    .map(function(d) { return d.trim().toLowerCase(); })
    .filter(function(d) { return d.length > 0; });
}

function shouldExclude_(email, excludeDomains) {
  var emailDomain = email.split('@')[1] || '';
  for (var i = 0; i < excludeDomains.length; i++) {
    if (emailDomain === excludeDomains[i] || email === excludeDomains[i]) return true;
  }
  return false;
}

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
  stopNudger_(true);

  var intervalHours = settings.triggerIntervalHours || 1;
  ScriptApp.newTrigger(TRIGGER_FUNCTION)
    .timeBased()
    .everyHours(intervalHours)
    .create();

  runFollowUpCheck();
  writeLog_('Nudger Started', '', '', 'Trigger created, running every ' +
    intervalHours + ' hour(s)');

  try {
    SpreadsheetApp.getUi().alert(
      '\u2705 Follow-Up Nudger is ACTIVE\n\n' +
      'Checking every ' + intervalHours + ' hour(s) for unreplied emails.\n' +
      'Follow-ups after: ' + settings.daysBeforeFollowUp + ' days\n' +
      'Auto-send: ' + (settings.autoSend === true || settings.autoSend === 'true' ? 'ON' : 'OFF') + '\n' +
      'Email digest: ' + (settings.sendDigest === true || settings.sendDigest === 'true' ? 'ON' : 'OFF') + '\n\n' +
      'To stop: \uD83D\uDD77 TAKScripts \u2192 Stop Follow-Up Nudger'
    );
  } catch(e) {
    Logger.log('\u2705 Follow-Up Nudger is ACTIVE');
  }
}

/**
 * Stops the Follow-Up Nudger and cleans up triggers.
 * @param {boolean} [silent] - If true, don't show UI alert.
 */
function stopNudger_(silent) {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  if (!silent) {
    writeLog_('Nudger Stopped', '', '', removed + ' trigger(s) removed');
    try {
      SpreadsheetApp.getUi().alert(
        '\u23F9 Follow-Up Nudger has been stopped.\n\n' +
        'No more follow-ups will be sent automatically.\n' +
        'Your dashboard data has been preserved.'
      );
    } catch(e) {
      Logger.log('\u23F9 Follow-Up Nudger has been stopped.');
    }
  }
}

/**
 * Public stop function (shows UI alert).
 */
function stopNudger() {
  stopNudger_(false);
}

/**
 * Displays test run results in a sidebar to avoid alert truncation.
 * @param {string} output - Plain-text output to display.
 */
function showTestResultsSidebar_(output) {
  var escaped = output
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
  var html = '<!DOCTYPE html><html><head><style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: Roboto, Arial, sans-serif; font-size: 12px; color: #333; background: #f9f9f9; }' +
    '.header { background: #1A1A1A; color: white; padding: 16px; text-align: center; }' +
    '.header h2 { font-size: 14px; font-weight: 600; color: #C9A84C; }' +
    '.content { padding: 16px; line-height: 1.6; }' +
    '</style></head><body>' +
    '<div class="header"><h2>\uD83E\uDDEA Test Run Results</h2></div>' +
    '<div class="content">' + escaped + '</div>' +
    '</body></html>';
  var panel = HtmlService.createHtmlOutput(html).setTitle('Test Run Results').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(panel);
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

    var searchDays = Math.max(SCAN_DAYS_BACK, daysThreshold + 7);
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
    try {
      showTestResultsSidebar_(output);
    } catch(e) {
      Logger.log('Sidebar display skipped: ' + e.message);
    }
    writeLog_('Test Run', '', '', pendingCount + ' pending follow-ups found');

  } catch (err) {
    try {
      SpreadsheetApp.getUi().alert('Error during test run: ' + err.message);
    } catch(e) {
      Logger.log('Error during test run: ' + err.message);
    }
    Logger.log('Test run error: ' + err.message);
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
  '    :root {' +
  '      --gold: #C9A84C;' +
  '      --gold-hover: #b8943c;' +
  '      --gold-glow: rgba(201,168,76,0.15);' +
  '      --bg-dark: #1A1A1A;' +
  '      --surface: #FAFAFA;' +
  '      --border: #E0E0E0;' +
  '    }' +
  '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
  '    body { font-family: "Segoe UI", system-ui, -apple-system, sans-serif; background: var(--surface); color: #1A1A1A; font-size: 13px; }' +
  '    .header { background: var(--bg-dark); color: white; padding: 20px 16px; text-align: center; }' +
  '    .header .logo { font-size: 24px; margin-bottom: 4px; }' +
  '    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }' +
  '    .header .brand { color: var(--gold); }' +
  '    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }' +
  '    .form { padding: 16px; }' +
  '    .section-title {' +
  '      font-size: 11px; font-weight: 700; text-transform: uppercase;' +
  '      letter-spacing: 1.2px; color: var(--gold); margin: 20px 0 12px;' +
  '      padding-bottom: 6px; border-bottom: 2px solid var(--gold);' +
  '    }' +
  '    .section-title:first-child { margin-top: 0; }' +
  '    .field { margin-bottom: 16px; }' +
  '    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }' +
  '    .field input, .field textarea, .field select {' +
  '      width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 6px;' +
  '      font-size: 13px; font-family: inherit; background: white; transition: border-color 0.2s, box-shadow 0.2s;' +
  '    }' +
  '    .field input:focus, .field textarea:focus, .field select:focus {' +
  '      outline: none; border-color: var(--gold); box-shadow: 0 0 0 3px var(--gold-glow);' +
  '    }' +
  '    .field textarea { min-height: 120px; resize: vertical; line-height: 1.5; }' +
  '    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }' +
  '    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }' +
  '    .toggle-wrap {' +
  '      display: flex; align-items: center; justify-content: space-between;' +
  '      padding: 12px; border: 1px solid #eee; border-radius: 8px;' +
  '      margin-bottom: 10px; background: #fff;' +
  '    }' +
  '    .toggle-wrap .label { font-size: 13px; font-weight: 500; }' +
  '    .toggle-wrap .sublabel { font-size: 11px; color: #999; margin-top: 2px; }' +
  '    .switch { position: relative; width: 44px; height: 24px; flex-shrink: 0; margin-left: 12px; }' +
  '    .switch input { opacity: 0; width: 0; height: 0; }' +
  '    .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background: #ccc; border-radius: 24px; transition: 0.3s; }' +
  '    .slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background: white; border-radius: 50%; transition: 0.3s; }' +
  '    .switch input:checked + .slider { background: var(--gold); }' +
  '    .switch input:checked + .slider:before { transform: translateX(20px); }' +
  '    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }' +
  '    .btn-primary { background: var(--gold); color: var(--bg-dark); border: 1px solid var(--gold); }' +
  '    .btn-primary:hover { background: var(--gold-hover); border-color: var(--gold-hover); }' +
  '    .btn-primary:disabled { opacity: 0.65; cursor: not-allowed; }' +
  '    .btn-secondary { background: white; color: #666; border: 1px solid var(--border); margin-top: 8px; }' +
  '    .btn-secondary:hover { border-color: #999; color: #333; }' +
  '    .status { text-align: center; padding: 10px 12px; font-size: 12px; font-weight: 500; margin-top: 10px; border-radius: 6px; display: none; animation: fadeIn 0.2s ease; }' +
  '    @keyframes fadeIn { from { opacity: 0; transform: translateY(-4px); } to { opacity: 1; transform: translateY(0); } }' +
  '    .status.success { display: block; background: #E6F4EA; color: #137333; border: 1px solid #ceead6; }' +
  '    .status.error { display: block; background: #FCE8E6; color: #C5221F; border: 1px solid #f5c6c2; }' +
  '    .divider { border: none; border-top: 1px solid rgba(201,168,76,0.15); margin: 20px 0; }' +
  '    .variables { background: #f5f5f5; border-radius: 6px; padding: 10px 12px; margin-top: 8px; }' +
  '    .variables code { display: inline-block; background: #e8e8e8; padding: 1px 6px; border-radius: 3px; font-size: 11px; margin: 2px; }' +
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
  '      <div><div class="label">Auto-Send Follow-Ups</div><div class="sublabel">Automatically send nudge emails</div></div>' +
  '      <label class="switch"><input type="checkbox" id="autoSend" /><span class="slider"></span></label>' +
  '    </div>' +
  '' +
  '    <div class="toggle-wrap">' +
  '      <div><div class="label">Email Digest</div><div class="sublabel">Get a summary of pending follow-ups</div></div>' +
  '      <label class="switch"><input type="checkbox" id="sendDigest" checked /><span class="slider"></span></label>' +
  '    </div>' +
  '' +
  '    <div class="toggle-wrap">' +
  '      <div><div class="label">Auto-Schedule</div><div class="sublabel">Run on a recurring trigger</div></div>' +
  '      <label class="switch"><input type="checkbox" id="autoSchedule" onchange="toggleIntervalField()" /><span class="slider"></span></label>' +
  '    </div>' +
  '' +
  '    <div class="field" id="intervalField" style="display: none;">' +
  '      <label>Run Every</label>' +
  '      <select id="triggerIntervalHours">' +
  '        <option value="1">Every hour</option>' +
  '        <option value="6">Every 6 hours</option>' +
  '        <option value="12">Every 12 hours</option>' +
  '        <option value="24">Once a day</option>' +
  '      </select>' +
  '    </div>' +
  '' +
  '    <div class="section-title">Urgency Thresholds</div>' +
  '' +
  '    <div class="field">' +
  '      <label>Overdue After (days)</label>' +
  '      <input type="number" id="overdueAfterDays" min="1" max="30" value="5" />' +
  '      <div class="help">Emails waiting this many days are marked Overdue</div>' +
  '    </div>' +
  '    <div class="field">' +
  '      <label>Urgent After (days)</label>' +
  '      <input type="number" id="urgentAfterDays" min="1" max="60" value="7" />' +
  '      <div class="help">Emails waiting this many days are marked Urgent</div>' +
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
  '      <label>Exclude Domains</label>' +
  '      <input type="text" id="excludeDomains" placeholder="noreply, newsletter, mailer-daemon" />' +
  '      <div class="help">Comma-separated. Emails from these domains will be skipped.</div>' +
  '    </div>' +
  '' +
  '    <div class="divider"></div>' +
  '' +
  '    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>' +
  '    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
  '    <div id="status" class="status"></div>' +
  '' +
  '  </div>' +
  '' +
  '  <script>' +
  '    google.script.run.withSuccessHandler(function(s) {' +
  '      document.getElementById("daysBeforeFollowUp").value = s.daysBeforeFollowUp || 3;' +
  '      document.getElementById("maxFollowUps").value = s.maxFollowUps || 2;' +
  '      document.getElementById("overdueAfterDays").value = s.overdueAfterDays || 5;' +
  '      document.getElementById("urgentAfterDays").value = s.urgentAfterDays || 7;' +
  '      document.getElementById("autoSend").checked = s.autoSend === true || s.autoSend === "true";' +
  '      document.getElementById("sendDigest").checked = s.sendDigest !== false && s.sendDigest !== "false";' +
  '      document.getElementById("followUpSubject").value = s.followUpSubject || "";' +
  '      document.getElementById("followUpMessage").value = s.followUpMessage || "";' +
  '      document.getElementById("excludeDomains").value = s.excludeDomains || "";' +
  '      document.getElementById("triggerIntervalHours").value = String(s.triggerIntervalHours || 1);' +
  '      var autoSchedule = s.autoSchedule === true || s.autoSchedule === "true";' +
  '      document.getElementById("autoSchedule").checked = autoSchedule;' +
  '      if (autoSchedule) document.getElementById("intervalField").style.display = "block";' +
  '    }).loadSettings();' +
  '' +
  '    function toggleIntervalField() {' +
  '      var checked = document.getElementById("autoSchedule").checked;' +
  '      document.getElementById("intervalField").style.display = checked ? "block" : "none";' +
  '    }' +
  '' +
  '    function save() {' +
  '      var settings = {' +
  '        daysBeforeFollowUp: parseInt(document.getElementById("daysBeforeFollowUp").value, 10) || 3,' +
  '        maxFollowUps: parseInt(document.getElementById("maxFollowUps").value, 10) || 2,' +
  '        overdueAfterDays: parseInt(document.getElementById("overdueAfterDays").value, 10) || 5,' +
  '        urgentAfterDays: parseInt(document.getElementById("urgentAfterDays").value, 10) || 7,' +
  '        autoSend: document.getElementById("autoSend").checked,' +
  '        sendDigest: document.getElementById("sendDigest").checked,' +
  '        autoSchedule: document.getElementById("autoSchedule").checked,' +
  '        triggerIntervalHours: parseInt(document.getElementById("triggerIntervalHours").value, 10) || 1,' +
  '        followUpSubject: document.getElementById("followUpSubject").value,' +
  '        followUpMessage: document.getElementById("followUpMessage").value,' +
  '        excludeDomains: document.getElementById("excludeDomains").value,' +
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

function showHelp() {
  var html = HtmlService.createHtmlOutput(getHelpHtml_())
    .setTitle('How to Use')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHelpHtml_() {
  return '<!DOCTYPE html>' +
  '<html><head><style>' +
  '* { box-sizing: border-box; margin: 0; padding: 0; }' +
  'body { font-family: Roboto, Arial, sans-serif; font-size: 13px; color: #333; background: #f9f9f9; }' +
  '.header { background: #1A1A1A; color: white; padding: 20px 16px 16px; text-align: center; }' +
  '.header .icon { font-size: 28px; margin-bottom: 6px; }' +
  '.header h2 { font-size: 15px; font-weight: 600; color: #C9A84C; margin-bottom: 2px; }' +
  '.header p { font-size: 11px; color: #888; }' +
  '.content { padding: 16px; }' +
  '.section { margin-bottom: 20px; }' +
  '.section h3 { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: #C9A84C; margin-bottom: 10px; padding-bottom: 6px; border-bottom: 1px solid #e0e0e0; }' +
  'ol, ul { padding-left: 18px; }' +
  'li { margin-bottom: 7px; line-height: 1.4; color: #444; font-size: 12px; }' +
  '.setting { margin-bottom: 10px; }' +
  '.setting strong { display: block; font-size: 12px; color: #1A1A1A; margin-bottom: 2px; }' +
  '.setting span { font-size: 11px; color: #666; line-height: 1.4; display: block; }' +
  '.tip { background: #FFF8E7; border-left: 3px solid #C9A84C; padding: 8px 10px; margin-bottom: 8px; border-radius: 0 4px 4px 0; font-size: 12px; color: #555; line-height: 1.4; }' +
  '.footer { text-align: center; padding: 12px; font-size: 10px; color: #aaa; border-top: 1px solid #eee; }' +
  '</style></head><body>' +
  '<div class="header"><div class="icon">\uD83D\uDD77</div>' +
  '<h2>Follow-Up Nudger</h2><p>Quick Reference Guide</p></div>' +
  '<div class="content">' +
  '<div class="section"><h3>Quick Start</h3><ol>' +
  '<li>Open <strong>\u2699\uFE0F Settings</strong> and set how many days before a follow-up is needed</li>' +
  '<li>Optionally configure urgency thresholds for Overdue and Urgent labels</li>' +
  '<li>Click <strong>\u25B6\uFE0F Start Follow-Up Nudger</strong> to activate</li>' +
  '<li>Emails with no reply appear in the dashboard with days waiting and a direct link</li>' +
  '<li>Use <strong>Snooze</strong> to hide a row for 7 days, or <strong>Mark Resolved</strong> to dismiss it</li>' +
  '</ol></div>' +
  '<div class="section"><h3>Settings Guide</h3>' +
  '<div class="setting"><strong>Days Before Follow-Up</strong><span>Emails with no reply after this many days get flagged</span></div>' +
  '<div class="setting"><strong>Auto-Send</strong><span>Automatically sends your follow-up template to flagged threads — use with care</span></div>' +
  '<div class="setting"><strong>Email Digest</strong><span>Sends you a summary email of all pending follow-ups on each run</span></div>' +
  '<div class="setting"><strong>Auto-Schedule</strong><span>Run automatically in the background on your chosen interval</span></div>' +
  '</div>' +
  '<div class="section"><h3>Tips</h3>' +
  '<div class="tip">Start with a <strong>3-day window</strong> to catch quick follow-ups on proposals and quotes</div>' +
  '<div class="tip">The dashboard shows <strong>Days Waiting</strong> — rows turn amber at Overdue and red at Urgent</div>' +
  '<div class="tip">Use <strong>Snooze</strong> on threads where you\'re waiting intentionally — they\'ll reappear automatically</div>' +
  '</div></div>' +
  '<div class="footer">TAKScripts \u00B7 takscripts.store</div>' +
  '</body></html>';
}
