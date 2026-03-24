/**
 * VIP Priority Alert by TAKScripts
 * ==================================
 * Monitors your inbox for emails from VIP contacts and sends
 * instant priority alerts so you never miss what matters.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - VIP contact list with priority levels (Critical, Important, Normal)
 * - Keyword triggers — flag emails by subject regardless of sender
 * - Alert methods: email notification with preview, auto-star + label
 * - Dashboard stats: Total Alerts · Critical · Important · Today · VIPs Tracked
 * - Quiet hours to prevent late-night notifications
 * - All alerts logged to a branded dashboard sheet
 * - Idempotent — tracks processed message IDs, safe to run multiple times
 * - Configurable check frequency (1–15 minutes)
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const SHEET_VIP = '⭐ VIP Contacts';
const DASHBOARD_SHEET_NAME = '📊 VIP Dashboard';
const DASHBOARD_HEADER_ROW = 3;
const LABEL_VIP = '⭐ VIP';
const PROP_SETTINGS = 'vip_settings';
const PROP_PROCESSED = 'vip_processed_ids';
const TRIGGER_FN = 'checkForVIPEmails';

const BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E8F5E9', successText: '#2E7D32',
  warningBg: '#FFF8E1', warningText: '#F57F17',
  errorBg: '#FFEBEE', errorText: '#C62828',
  headerFont: 'Roboto Mono', bodyFont: 'Roboto',
};

const LOG_HEADERS = ['Timestamp', 'Sender', 'Subject', 'Priority', 'Action Taken', 'Trigger'];

// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🕷 TAKScripts')
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('▶️ Start Monitoring', 'startMonitoring')
    .addItem('⏹ Stop Monitoring', 'stopMonitoring')
    .addSeparator()
    .addItem('🧪 Test Run', 'testRun')
    .addItem('📊 View Dashboard', 'viewDashboard')
    .addItem('🔄 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('VIP Priority Alert — Settings')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function viewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME) || getOrCreateDashboard_();
  ss.setActiveSheet(sheet);
}

function showAbout() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Segoe UI', system-ui, sans-serif; padding: 20px; text-align: center;">
      <div style="font-size: 32px; margin-bottom: 8px;">🕷</div>
      <h2 style="margin: 0 0 4px; font-size: 18px;">VIP Priority Alert</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 · by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
        Never miss an email from the people who matter most.<br>
        Part of the <strong>TAKScripts</strong> collection.
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


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_SETTINGS, JSON.stringify(settings));
  return { success: true };
}

function getDefaultSettings_() {
  return {
    checkFrequency: '5',
    alertEmail: Session.getActiveUser().getEmail() || '',
    keywords: 'urgent, asap, time-sensitive, action required',
    quietStart: '22:00',
    quietEnd: '07:00',
    quietEnabled: false,
  };
}

function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_SETTINGS);
  if (!raw) return getDefaultSettings_();
  const saved = JSON.parse(raw);
  const defaults = getDefaultSettings_();
  for (const key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}


// ═══════════════════════════════════════════
// SHEET SETUP
// ═══════════════════════════════════════════

/**
 * Ensures the VIP Contacts sheet exists with proper formatting.
 * This is the user's config sheet — no dashboard, just a clean list.
 */
function ensureVIPSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_VIP);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_VIP, 0);
    const headers = ['Name', 'Email', 'Priority', 'Alert Method', 'Notes'];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontFamily(BRAND.headerFont)
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground(BRAND.darkBg)
      .setFontColor(BRAND.gold)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 32);
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 220);

    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['🔴 Critical', '🟡 Important', '🟢 Normal'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 3, 200, 1).setDataValidation(priorityRule);

    const methodRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['📧 Email Alert', '⭐ Star + Label Only'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, 200, 1).setDataValidation(methodRule);

    sheet.getRange(2, 1, 200, headers.length)
      .setFontFamily(BRAND.bodyFont)
      .setFontSize(10);

    sheet.appendRow([
      'Jane Doe (example)',
      'jane@example.com',
      '🔴 Critical',
      '📧 Email Alert',
      'CEO — always alert',
    ]);
  }

  return sheet;
}

/**
 * Gets or creates the VIP Dashboard sheet with stats bar and alert log headers.
 * Idempotent — safe to call multiple times.
 */
function getOrCreateDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    // Migration: rename old alert log sheet if it exists
    const oldSheet = ss.getSheetByName('📋 Alert Log');
    if (oldSheet) {
      oldSheet.setName(DASHBOARD_SHEET_NAME);
      oldSheet.clear();
      sheet = oldSheet;
    } else {
      sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
    }
  } else {
    const headerCheck = sheet.getRange(DASHBOARD_HEADER_ROW, 1).getValue();
    if (headerCheck === LOG_HEADERS[0]) {
      return sheet; // Already set up
    }
  }

  const numCols = LOG_HEADERS.length;

  // ── Row 1: Stat values ──────────────────────────────────
  sheet.getRange(1, 1, 1, numCols)
    .setValues([['—', '—', '—', '—', '—', '']])
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
    .setValues([['TOTAL ALERTS', 'CRITICAL', 'IMPORTANT', 'TODAY', 'VIPS TRACKED', '']])
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
    .setValues([LOG_HEADERS])
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
  sheet.setColumnWidth(1, 170); // Timestamp
  sheet.setColumnWidth(2, 240); // Sender
  sheet.setColumnWidth(3, 300); // Subject
  sheet.setColumnWidth(4, 140); // Priority
  sheet.setColumnWidth(5, 180); // Action Taken
  sheet.setColumnWidth(6, 160); // Trigger

  return sheet;
}

/**
 * Recalculates all stats from alert log data and VIP contacts, updates stats bar.
 */
function refreshDashboardStats() {
  const sheet = getOrCreateDashboard_();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const lastRow = sheet.getLastRow();

  let totalAlerts = 0;
  let critical = 0;
  let important = 0;
  let today = 0;

  if (lastRow >= dataStartRow) {
    const numRows = lastRow - dataStartRow + 1;
    const data = sheet.getRange(dataStartRow, 1, numRows, LOG_HEADERS.length).getValues();
    const now = new Date();

    for (const row of data) {
      if (!row[0]) continue; // Skip empty rows
      totalAlerts++;
      const priority = (row[3] || '').toString();
      if (priority.includes('Critical')) critical++;
      else if (priority.includes('Important')) important++;

      const ts = row[0];
      if (ts instanceof Date &&
          ts.getDate() === now.getDate() &&
          ts.getMonth() === now.getMonth() &&
          ts.getFullYear() === now.getFullYear()) {
        today++;
      }
    }
  }

  // Count VIPs from VIP Contacts sheet
  const vips = getVIPList_();

  sheet.getRange(1, 1, 1, 5).setValues([[
    totalAlerts || '—',
    critical || '—',
    important || '—',
    today || '—',
    vips.length || '—',
  ]]);
}


// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

function checkForVIPEmails() {
  const settings = loadSettings();
  const vips = getVIPList_();

  if (vips.length === 0) {
    Logger.log('No VIP contacts configured. Add contacts to the ⭐ VIP Contacts sheet.');
    return;
  }

  if (settings.quietEnabled && isQuietHours_(settings.quietStart, settings.quietEnd)) {
    Logger.log('Quiet hours active. Skipping check.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_PROCESSED) || '[]';
  const processedArr = JSON.parse(processedRaw);
  const processed = new Set(processedArr);

  const keywords = (settings.keywords || '')
    .split(',')
    .map(k => k.trim().toLowerCase())
    .filter(Boolean);

  const threads = GmailApp.search('is:unread is:inbox newer_than:1d', 0, 50);
  let alertCount = 0;
  const newProcessed = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    const msgId = msg.getId();

    if (processed.has(msgId)) continue;

    const senderEmail = extractEmail_(msg.getFrom());
    const senderName = extractName_(msg.getFrom());
    const subject = msg.getSubject() || '(no subject)';
    const snippet = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ').trim();

    const vipMatch = vips.find(v => v.email === senderEmail);
    const subjectLower = subject.toLowerCase();
    const keywordHit = keywords.find(kw => subjectLower.includes(kw));

    if (!vipMatch && !keywordHit) continue;

    const priority = vipMatch ? vipMatch.priority : '🟡 Important';
    const alertMethod = vipMatch ? vipMatch.alertMethod : '📧 Email Alert';
    const triggerReason = vipMatch
      ? 'VIP: ' + (vipMatch.name || senderEmail)
      : 'Keyword: "' + keywordHit + '"';

    applyVIPLabel_(thread);
    thread.getMessages().forEach(m => m.star());

    if (alertMethod === '📧 Email Alert') {
      sendAlertEmail_(settings, senderName, senderEmail, subject, snippet, priority, triggerReason);
    }

    logAlert_(senderEmail, subject, priority,
      alertMethod === '📧 Email Alert' ? 'Email alert sent' : 'Starred + labeled',
      triggerReason);

    newProcessed.push(msgId);
    alertCount++;
    Logger.log('VIP alert: ' + senderEmail + ' — ' + subject);
  }

  const allProcessed = [...processedArr, ...newProcessed].slice(-500);
  props.setProperty(PROP_PROCESSED, JSON.stringify(allProcessed));

  Logger.log('Done. Triggered ' + alertCount + ' VIP alerts.');
}


// ═══════════════════════════════════════════
// VIP LIST MANAGEMENT
// ═══════════════════════════════════════════

function getVIPList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_VIP);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const vips = [];

  for (const row of data) {
    const email = (row[1] || '').toString().trim().toLowerCase();
    if (!email || !email.includes('@')) continue;

    vips.push({
      name: (row[0] || '').toString().trim(),
      email: email,
      priority: (row[2] || '🟢 Normal').toString().trim(),
      alertMethod: (row[3] || '📧 Email Alert').toString().trim(),
    });
  }

  return vips;
}


// ═══════════════════════════════════════════
// GMAIL HELPERS
// ═══════════════════════════════════════════

function applyVIPLabel_(thread) {
  let label = GmailApp.getUserLabelByName(LABEL_VIP);
  if (!label) {
    label = GmailApp.createLabel(LABEL_VIP);
  }
  label.addToThread(thread);
}

function sendAlertEmail_(settings, senderName, senderEmail, subject, snippet, priority, triggerReason) {
  const alertTo = settings.alertEmail || Session.getActiveUser().getEmail();
  if (!alertTo) return;

  const priorityColor = priority.includes('Critical') ? '#C62828'
    : priority.includes('Important') ? '#F57F17'
    : '#2E7D32';

  const priorityBg = priority.includes('Critical') ? '#FFEBEE'
    : priority.includes('Important') ? '#FFF8E1'
    : '#E8F5E9';

  const htmlBody = `
    <div style="font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; max-width: 560px; margin: 0 auto;">
      <div style="background: #1A1A1A; padding: 20px 24px; border-radius: 8px 8px 0 0; text-align: center;">
        <span style="font-size: 22px;">🕷</span>
        <span style="color: #C9A84C; font-size: 16px; font-weight: 700; letter-spacing: 0.5px; vertical-align: middle; margin-left: 8px;">VIP Priority Alert</span>
      </div>
      <div style="background: #FFFFFF; padding: 24px; border: 1px solid #E0E0E0; border-top: none;">
        <div style="background: ${priorityBg}; color: ${priorityColor}; padding: 10px 14px; border-radius: 6px; font-weight: 600; font-size: 13px; margin-bottom: 16px;">
          ${priority} — ${triggerReason}
        </div>
        <table style="width: 100%; font-size: 13px; color: #333; border-collapse: collapse;">
          <tr>
            <td style="padding: 8px 0; color: #888; width: 70px; vertical-align: top;">From</td>
            <td style="padding: 8px 0; font-weight: 600;">${senderName ? senderName + ' &lt;' + senderEmail + '&gt;' : senderEmail}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #888; vertical-align: top;">Subject</td>
            <td style="padding: 8px 0; font-weight: 600;">${escapeHtml_(subject)}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #888; vertical-align: top;">Preview</td>
            <td style="padding: 8px 0; color: #555; line-height: 1.5;">${escapeHtml_(snippet)}${snippet.length >= 200 ? '...' : ''}</td>
          </tr>
        </table>
        <div style="margin-top: 20px; text-align: center;">
          <a href="https://mail.google.com" target="_blank"
             style="display: inline-block; background: #1A1A1A; color: #C9A84C; padding: 10px 28px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 13px; letter-spacing: 0.3px;">
            Open Gmail →
          </a>
        </div>
      </div>
      <div style="background: #FAFAFA; padding: 14px 24px; border: 1px solid #E0E0E0; border-top: none; border-radius: 0 0 8px 8px; text-align: center;">
        <span style="font-size: 11px; color: #999;">Sent by <strong>VIP Priority Alert</strong> via <a href="https://takscripts.store" style="color: #C9A84C; text-decoration: none;">TAKScripts</a></span>
      </div>
    </div>
  `;

  GmailApp.sendEmail(alertTo, '⭐ VIP Alert: ' + subject + ' — from ' + (senderName || senderEmail), '', {
    htmlBody: htmlBody,
    name: 'VIP Priority Alert',
  });
}


// ═══════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════

/**
 * Logs an alert to the VIP Dashboard sheet and refreshes stats.
 */
function logAlert_(sender, subject, priority, action, trigger) {
  const sheet = getOrCreateDashboard_();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const newRow = Math.max(sheet.getLastRow() + 1, dataStartRow);

  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setValues([
    [new Date(), sender, subject, priority, action, trigger],
  ]);

  // Base row styling
  const isEven = (newRow - DASHBOARD_HEADER_ROW) % 2 === 0;
  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setBackground(isEven ? BRAND.lightGray : BRAND.white);
  sheet.setRowHeight(newRow, 30);

  // Timestamp format
  sheet.getRange(newRow, 1).setNumberFormat('MMM d, yyyy h:mm a');

  // Priority cell color
  const priorityCell = sheet.getRange(newRow, 4);
  if (priority.includes('Critical')) {
    priorityCell.setBackground(BRAND.errorBg).setFontColor(BRAND.errorText).setFontWeight('bold');
  } else if (priority.includes('Important')) {
    priorityCell.setBackground(BRAND.warningBg).setFontColor(BRAND.warningText).setFontWeight('bold');
  } else {
    priorityCell.setBackground(BRAND.successBg).setFontColor(BRAND.successText).setFontWeight('bold');
  }

  // Action cell color
  sheet.getRange(newRow, 5)
    .setBackground(BRAND.successBg)
    .setFontColor(BRAND.successText);

  refreshDashboardStats();
}


// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

function extractEmail_(fromField) {
  const match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

function extractName_(fromField) {
  const match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim() : '';
}

function escapeHtml_(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function isQuietHours_(startStr, endStr) {
  if (!startStr || !endStr) return false;

  const now = new Date();
  const currentMinutes = now.getHours() * 60 + now.getMinutes();

  const [startH, startM] = startStr.split(':').map(Number);
  const [endH, endM] = endStr.split(':').map(Number);
  const startMinutes = startH * 60 + startM;
  const endMinutes = endH * 60 + endM;

  if (startMinutes > endMinutes) {
    return currentMinutes >= startMinutes || currentMinutes < endMinutes;
  }
  return currentMinutes >= startMinutes && currentMinutes < endMinutes;
}


// ═══════════════════════════════════════════
// CONTROLS
// ═══════════════════════════════════════════

function startMonitoring() {
  const settings = loadSettings();
  ensureVIPSheet();
  getOrCreateDashboard_();

  stopMonitoring_(false);

  const freq = parseInt(settings.checkFrequency, 10) || 5;

  ScriptApp.newTrigger(TRIGGER_FN)
    .timeBased()
    .everyMinutes(freq)
    .create();

  checkForVIPEmails();

  const vips = getVIPList_();
  try {
    SpreadsheetApp.getUi().alert(
      '✅ VIP Monitoring is ACTIVE\n\n' +
      'Tracking ' + vips.length + ' VIP contact(s).\n' +
      'Checking every ' + freq + ' minute(s).\n\n' +
      'Add or edit VIPs in the "⭐ VIP Contacts" sheet.\n' +
      'To stop: 🕷 TAKScripts → ⏹ Stop Monitoring'
    );
  } catch(e) {
    Logger.log('✅ VIP Monitoring is ACTIVE\n\nTracking ' + vips.length + ' VIP contact(s).\nChecking every ' + freq + ' minute(s).\n\nAdd or edit VIPs in the "⭐ VIP Contacts" sheet.\nTo stop: 🕷 TAKScripts → ⏹ Stop Monitoring');
  }
}

function stopMonitoring() {
  stopMonitoring_(true);
}

function stopMonitoring_(showAlert) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FN) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  if (showAlert) {
    PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
    try {
      SpreadsheetApp.getUi().alert(
        '⏹ VIP Monitoring STOPPED\n\n' +
        'No more alerts will be sent.\n' +
        'Processed message history has been cleared.\n\n' +
        'To resume: 🕷 TAKScripts → ▶️ Start Monitoring'
      );
    } catch(e) {
      Logger.log('⏹ VIP Monitoring STOPPED\n\nNo more alerts will be sent.\nProcessed message history has been cleared.\n\nTo resume: 🕷 TAKScripts → ▶️ Start Monitoring');
    }
  }
}

function testRun() {
  ensureVIPSheet();
  const vips = getVIPList_();
  const settings = loadSettings();

  const keywords = (settings.keywords || '')
    .split(',')
    .map(k => k.trim().toLowerCase())
    .filter(Boolean);

  const threads = GmailApp.search('is:inbox newer_than:1d', 0, 10);
  const results = [];

  results.push('🧪 TEST RUN — No alerts will be sent');
  results.push('VIPs loaded: ' + vips.length + ' contact(s)');
  results.push('Keywords: ' + (keywords.length > 0 ? keywords.join(', ') : '(none)'));
  results.push('Checked: ' + threads.length + ' recent thread(s)');
  results.push('―――――――――――――――――――――――――');

  let matchCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    const senderEmail = extractEmail_(msg.getFrom());
    const senderName = extractName_(msg.getFrom());
    const subject = msg.getSubject() || '(no subject)';

    const vipMatch = vips.find(v => v.email === senderEmail);
    const subjectLower = subject.toLowerCase();
    const keywordHit = keywords.find(kw => subjectLower.includes(kw));

    const displayFrom = senderName ? senderName + ' <' + senderEmail + '>' : senderEmail;

    if (vipMatch || keywordHit) {
      const priority = vipMatch ? vipMatch.priority : '🟡 Important';
      const reason = vipMatch ? 'VIP match' : 'Keyword: "' + keywordHit + '"';
      results.push('✅ ALERT — ' + priority);
      results.push('   From: ' + displayFrom);
      results.push('   Subject: ' + subject);
      results.push('   Reason: ' + reason);
      matchCount++;
    } else {
      results.push('⬜ No match');
      results.push('   From: ' + displayFrom);
      results.push('   Subject: ' + subject);
    }
    results.push('');
  }

  if (threads.length === 0) {
    results.push('No recent emails found.');
  } else {
    results.push('―――――――――――――――――――――――――');
    results.push(matchCount + ' of ' + threads.length + ' emails would trigger alerts.');
  }

  const output = results.join('\n');
  Logger.log(output);
  try {
    SpreadsheetApp.getUi().alert(output);
  } catch(e) {
    Logger.log(output);
  }
}

/**
 * One-time setup — creates both sheets with proper formatting.
 */
function install() {
  ensureVIPSheet();
  getOrCreateDashboard_();
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Setup Complete\n\n' +
      'Sheets created:\n' +
      '• ⭐ VIP Contacts — add your VIPs here\n' +
      '• 📊 VIP Dashboard — alerts will appear here\n\n' +
      'Next: Open 🕷 TAKScripts → ⚙️ Settings to configure.'
    );
  } catch(e) {
    Logger.log('✅ Setup Complete\n\nSheets created:\n• ⭐ VIP Contacts — add your VIPs here\n• 📊 VIP Dashboard — alerts will appear here\n\nNext: Open 🕷 TAKScripts → ⚙️ Settings to configure.');
  }
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
    body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: #FAFAFA; color: #1a1a1a; font-size: 13px; }
    .header { background: #1A1A1A; color: white; padding: 20px 16px; text-align: center; }
    .header .logo { font-size: 24px; margin-bottom: 4px; }
    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }
    .header .brand { color: #C9A84C; }
    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }
    .form { padding: 16px; }
    .section-title {
      font-size: 11px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 1.2px; color: #C9A84C; margin: 20px 0 12px;
      padding-bottom: 6px; border-bottom: 1px solid #eee;
    }
    .section-title:first-child { margin-top: 0; }
    .field { margin-bottom: 16px; }
    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }
    .field input, .field textarea, .field select {
      width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px;
      font-size: 13px; font-family: inherit; background: white; transition: border-color 0.2s;
    }
    .field input:focus, .field textarea:focus, .field select:focus {
      outline: none; border-color: #C9A84C; box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea { min-height: 80px; resize: vertical; line-height: 1.5; }
    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .toggle-row { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
    .toggle-row input[type="checkbox"] { width: 18px; height: 18px; accent-color: #C9A84C; cursor: pointer; }
    .toggle-row label { font-size: 13px; font-weight: 500; color: #333; cursor: pointer; text-transform: none; letter-spacing: 0; }
    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }
    .btn-primary { background: #1A1A1A; color: #C9A84C; border: 1px solid #C9A84C; }
    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }
    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .status { text-align: center; padding: 8px; font-size: 12px; margin-top: 8px; border-radius: 6px; display: none; }
    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }
    .status.error { display: block; background: #FFEBEE; color: #C62828; }
    .divider { border-top: 1px solid #eee; margin: 20px 0; }
    .info-box { background: #FFF8E1; border-radius: 6px; padding: 10px 12px; font-size: 12px; color: #F57F17; line-height: 1.5; margin-top: 8px; }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo">🕷</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">VIP Priority Alert · Settings</div>
  </div>

  <div class="form">
    <div class="section-title">Monitoring</div>

    <div class="field">
      <label>Check Frequency</label>
      <select id="checkFrequency">
        <option value="1">Every 1 minute</option>
        <option value="5" selected>Every 5 minutes</option>
        <option value="10">Every 10 minutes</option>
        <option value="15">Every 15 minutes</option>
      </select>
      <div class="help">How often to scan for new VIP emails.</div>
    </div>

    <div class="field">
      <label>Alert Email Address</label>
      <input type="email" id="alertEmail" placeholder="you@example.com" />
      <div class="help">Where to send VIP alert notifications.</div>
    </div>

    <div class="section-title">Keyword Triggers</div>

    <div class="field">
      <label>Subject Keywords</label>
      <textarea id="keywords" placeholder="urgent, asap, time-sensitive"></textarea>
      <div class="help">Comma-separated. Emails with these words in the subject will trigger an alert regardless of sender.</div>
    </div>

    <div class="section-title">Quiet Hours</div>

    <div class="toggle-row">
      <input type="checkbox" id="quietEnabled" />
      <label for="quietEnabled">Enable quiet hours</label>
    </div>

    <div class="row">
      <div class="field">
        <label>Start</label>
        <input type="time" id="quietStart" value="22:00" />
      </div>
      <div class="field">
        <label>End</label>
        <input type="time" id="quietEnd" value="07:00" />
      </div>
    </div>

    <div class="info-box">
      During quiet hours, VIP emails will still be starred and labeled, but no email alerts will be sent.
    </div>

    <div class="divider"></div>

    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    google.script.run.withSuccessHandler(function(s) {
      document.getElementById('checkFrequency').value = s.checkFrequency || '5';
      document.getElementById('alertEmail').value = s.alertEmail || '';
      document.getElementById('keywords').value = s.keywords || '';
      document.getElementById('quietStart').value = s.quietStart || '22:00';
      document.getElementById('quietEnd').value = s.quietEnd || '07:00';
      document.getElementById('quietEnabled').checked = !!s.quietEnabled;
    }).loadSettings();

    function save() {
      const settings = {
        checkFrequency: document.getElementById('checkFrequency').value,
        alertEmail: document.getElementById('alertEmail').value,
        keywords: document.getElementById('keywords').value,
        quietStart: document.getElementById('quietStart').value,
        quietEnd: document.getElementById('quietEnd').value,
        quietEnabled: document.getElementById('quietEnabled').checked,
      };

      const freq = parseInt(settings.checkFrequency, 10);
      const statusEl = document.getElementById('status');
      if (freq < 1 || freq > 15) {
        statusEl.textContent = '✕ Check frequency must be between 1 and 15 minutes.';
        statusEl.className = 'status error';
        return;
      }

      const saveBtn = document.getElementById('saveBtn');
      saveBtn.disabled = true;
      saveBtn.textContent = 'Saving…';
      google.script.run
        .withSuccessHandler(function() {
          statusEl.textContent = '✓ Settings saved successfully';
          statusEl.className = 'status success';
          saveBtn.textContent = '✓ Saved!';
          setTimeout(function() {
            saveBtn.textContent = 'Save Settings';
            saveBtn.disabled = false;
          }, 2500);
        })
        .withFailureHandler(function(err) {
          statusEl.textContent = '✕ Error: ' + err.message;
          statusEl.className = 'status error';
          saveBtn.textContent = 'Save Settings';
          saveBtn.disabled = false;
        })
        .saveSettings(settings);
    }
  </script>
</body>
</html>`;
}
