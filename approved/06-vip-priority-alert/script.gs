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
const DASHBOARD_HEADER_ROW = 4;
const LABEL_VIP = '⭐ VIP';
const PROP_SETTINGS = 'vip_settings';
const PROP_PROCESSED = 'vip_processed_ids';
const TRIGGER_FN = 'checkForVIPEmails';

const BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E6F4EA', successText: '#137333',
  warningBg: '#FEF7E0', warningText: '#B06000',
  errorBg: '#FCE8E6', errorText: '#C5221F',
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
    .addItem('❓ How to Use', 'showHelp')
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('VIP Priority Alert — Settings')
    .setWidth(430);
  SpreadsheetApp.getUi().showSidebar(html);
}

function viewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME) || getOrCreateDashboard_();
  ss.setActiveSheet(sheet);
  refreshDashboardStats();
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
    lookbackCount: 50,
    excludeDomains: 'noreply, no-reply, mailer-daemon, notifications',
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
    const headers = ['Name', 'Email', 'Priority', 'Alert Method', 'Notes', 'Last Alerted'];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontFamily(BRAND.headerFont)
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground(BRAND.darkBg)
      .setFontColor(BRAND.gold)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(3, 130); // Priority — center-aligned
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 220);
    sheet.setColumnWidth(6, 160); // Last Alerted

    // Priority — center-align and data validation (rows 2–50 only)
    sheet.getRange(2, 3, 49, 1).setHorizontalAlignment('center');
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['🔴 Critical', '🟡 Important', '🟢 Normal'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 3, 49, 1).setDataValidation(priorityRule);

    const methodRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['📧 Email Alert', '⭐ Star + Label Only'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, 49, 1).setDataValidation(methodRule);

    sheet.getRange(2, 1, 49, headers.length)
      .setFontFamily(BRAND.bodyFont)
      .setFontSize(10);

    // Last Alerted column — timestamp format
    sheet.getRange(2, 6, 49, 1).setNumberFormat('MMM d, yyyy h:mm a');

    // Alternating row tint via banding
    sheet.getRange(2, 1, 49, headers.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstBandColor('#FFFFFF')
      .setSecondBandColor('#F8F8F8');

    sheet.appendRow([
      'Jane Doe (example)',
      'jane@example.com',
      '🔴 Critical',
      '📧 Email Alert',
      'CEO — always alert',
      '',
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

  // ── Row 1: Title bar ────────────────────────────────────
  sheet.getRange(1, 1, 1, numCols)
    .merge()
    .setValue('📊  VIP DASHBOARD')
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
    .setValues([['—', '—', '—', '—', '—', '']])
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
    .setValues([['TOTAL ALERTS', 'CRITICAL', 'IMPORTANT', 'TODAY', 'VIPS TRACKED', '']])
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
    .setValues([LOG_HEADERS])
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
  sheet.setColumnWidth(1, 170); // Timestamp
  sheet.setColumnWidth(2, 240); // Sender
  sheet.setColumnWidth(3, 300); // Subject
  sheet.setColumnWidth(4, 140); // Priority
  sheet.setColumnWidth(5, 180); // Action Taken
  sheet.setColumnWidth(6, 160); // Trigger

  // ── Priority-row conditional formatting ─────────────────
  const firstDataRow = DASHBOARD_HEADER_ROW + 1;
  const cfRange = sheet.getRange(firstDataRow, 1, 1000, numCols);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
        [`=ISNUMBER(SEARCH("Critical",$D${firstDataRow}))`])
      .setBackground('#FCE4EC')
      .setRanges([cfRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
        [`=ISNUMBER(SEARCH("Important",$D${firstDataRow}))`])
      .setBackground('#FFF8E1')
      .setRanges([cfRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
        [`=ISNUMBER(SEARCH("Normal",$D${firstDataRow}))`])
      .setBackground('#E8F5E9')
      .setRanges([cfRange])
      .build(),
  ]);

  // ── Hide gridlines (whole spreadsheet) ──────────────────
  ss.setHiddenGridlines(true);

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

  sheet.getRange(2, 1, 1, 5).setValues([[
    totalAlerts || '—',
    critical || '—',
    important || '—',
    today || '—',
    vips.length || '—',
  ]]);

  var lastRunRaw = PropertiesService.getScriptProperties().getProperty('vip_last_run');
  if (lastRunRaw) {
    var lastRun = Utilities.formatDate(new Date(lastRunRaw), Session.getScriptTimeZone(), 'MMM d, h:mm a');
    sheet.getRange(2, 6).setValue(lastRun);
    sheet.getRange(3, 6).setValue('LAST RUN');
  }
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

  const threads = GmailApp.search('is:unread is:inbox newer_than:1d', 0, settings.lookbackCount || 50);
  let alertCount = 0;
  const newProcessed = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    const msgId = msg.getId();

    if (processed.has(msgId)) continue;

    const rawFrom = msg.getFrom();
    const rawFromLower = rawFrom.toLowerCase();
    const senderEmail = extractEmail_(rawFromLower);
    const senderName = extractName_(rawFromLower);
    const subject = msg.getSubject() || '(no subject)';
    const snippet = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ').trim();

    const vipMatch = vips.find(v => v.email === senderEmail);
    const subjectLower = subject.toLowerCase();
    const senderDomain = senderEmail.split('@')[1] || '';
    const excludeList = (settings.excludeDomains || '').split(',').map(function(d) { return d.trim().toLowerCase(); });
    const isDomainExcluded = excludeList.some(function(d) { return d && senderDomain.indexOf(d) !== -1; });
    const keywordHit = isDomainExcluded ? null : keywords.find(kw => subjectLower.includes(kw));

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
  PropertiesService.getScriptProperties().setProperty('vip_last_run', new Date().toISOString());

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

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const email = (row[1] || '').toString().trim().toLowerCase();
    if (!email || !email.includes('@')) {
      console.log('VIP list: skipping row ' + (i + 2) + ' — invalid or missing email: "' + (row[1] || '') + '"');
      continue;
    }

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

  const priorityColor = priority.includes('Critical') ? BRAND.errorText
    : priority.includes('Important') ? BRAND.warningText
    : BRAND.successText;

  const priorityBg = priority.includes('Critical') ? BRAND.errorBg
    : priority.includes('Important') ? BRAND.warningBg
    : BRAND.successBg;

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

  const now = new Date();
  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setValues([
    [now, sender, subject, priority, action, trigger],
  ]);

  // Row styling
  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10);
  sheet.setRowHeight(newRow, 30);

  // Timestamp format
  sheet.getRange(newRow, 1).setNumberFormat('MMM d, yyyy h:mm a');

  // Priority cell — bold + text color (background handled by conditional formatting)
  const priorityCell = sheet.getRange(newRow, 4);
  if (priority.includes('Critical')) {
    priorityCell.setFontColor(BRAND.errorText).setFontWeight('bold');
  } else if (priority.includes('Important')) {
    priorityCell.setFontColor(BRAND.warningText).setFontWeight('bold');
  } else {
    priorityCell.setFontColor(BRAND.successText).setFontWeight('bold');
  }

  // Action cell — success tint
  sheet.getRange(newRow, 5)
    .setBackground(BRAND.successBg)
    .setFontColor(BRAND.successText);

  // Update Last Alerted in VIP Contacts sheet
  try {
    const vipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIP);
    if (vipSheet && vipSheet.getLastRow() > 1) {
      const emailData = vipSheet.getRange(2, 2, vipSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < emailData.length; i++) {
        if ((emailData[i][0] || '').toString().trim().toLowerCase() === sender.toLowerCase()) {
          vipSheet.getRange(i + 2, 6).setValue(now).setNumberFormat('MMM d, yyyy h:mm a');
          break;
        }
      }
    }
  } catch(e) {
    Logger.log('Could not update Last Alerted: ' + e.message);
  }

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

  // Remove any existing triggers for this function first
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === TRIGGER_FN) {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }

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
    const rawFrom = msg.getFrom();
    const rawFromLower = rawFrom.toLowerCase();
    const senderEmail = extractEmail_(rawFromLower);
    const senderName = extractName_(rawFromLower);
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
    showTestResultsSidebar_(output);
  } catch(e) {
    Logger.log('Sidebar display skipped: ' + e.message);
  }
}

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
    '.content { padding: 16px; line-height: 1.7; }' +
    '</style></head><body>' +
    '<div class="header"><h2>🧪 Test Run Results</h2></div>' +
    '<div class="content">' + escaped + '</div>' +
    '</body></html>';
  var panel = HtmlService.createHtmlOutput(html).setTitle('Test Run Results').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(panel);
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
    :root {
      --gold: #C9A84C;
      --gold-hover: #b8943c;
      --gold-glow: rgba(201,168,76,0.15);
      --bg-dark: #1A1A1A;
      --surface: #FAFAFA;
      --border: #E0E0E0;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--surface); color: #1a1a1a; font-size: 13px; }
    .header { background: var(--bg-dark); color: white; padding: 20px 16px; text-align: center; }
    .header .logo { font-size: 24px; margin-bottom: 4px; }
    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }
    .header .brand { color: var(--gold); }
    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }
    .form { padding: 16px; }
    .section-title {
      font-size: 11px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 1.2px; color: var(--gold); margin: 20px 0 12px;
      padding-bottom: 6px; border-bottom: 2px solid var(--gold);
    }
    .section-title:first-child { margin-top: 0; }
    .field { margin-bottom: 16px; }
    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }
    .field input, .field textarea, .field select {
      width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 6px;
      font-size: 13px; font-family: inherit; background: white; transition: border-color 0.2s, box-shadow 0.2s;
    }
    .field input:focus, .field textarea:focus, .field select:focus {
      outline: none; border-color: var(--gold); box-shadow: 0 0 0 3px var(--gold-glow);
    }
    .field textarea { min-height: 80px; resize: vertical; line-height: 1.5; }
    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .toggle-row { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
    .toggle-row input[type="checkbox"] { width: 18px; height: 18px; accent-color: var(--gold); cursor: pointer; }
    .toggle-row label { font-size: 13px; font-weight: 500; color: #333; cursor: pointer; text-transform: none; letter-spacing: 0; }
    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }
    .btn-primary { background: var(--gold); color: var(--bg-dark); border: 1px solid var(--gold); }
    .btn-primary:hover { background: var(--gold-hover); border-color: var(--gold-hover); }
    .btn-primary:disabled { opacity: 0.65; cursor: not-allowed; }
    .btn-secondary { background: white; color: #666; border: 1px solid var(--border); margin-top: 8px; }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .status {
      text-align: center; padding: 10px 12px; font-size: 12px; font-weight: 500;
      margin-top: 10px; border-radius: 6px; display: none;
      animation: fadeIn 0.2s ease;
    }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(-4px); } to { opacity: 1; transform: translateY(0); } }
    .status.success { display: block; background: #E6F4EA; color: #137333; border: 1px solid #ceead6; }
    .status.error { display: block; background: #FCE8E6; color: #C5221F; border: 1px solid #f5c6c2; }
    .divider { border: none; border-top: 1px solid var(--gold-glow); margin: 20px 0; }
    .info-box { background: #FEF7E0; border-left: 3px solid var(--gold); border-radius: 0 6px 6px 0; padding: 10px 12px; font-size: 12px; color: #B06000; line-height: 1.5; margin-top: 8px; }
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

    <div class="field">
      <label>Threads to Scan</label>
      <input type="number" id="lookbackCount" min="10" max="500" value="50" />
      <div class="help">Number of recent email threads to check on each run</div>
    </div>

    <div class="section-title">Keyword Triggers</div>

    <div class="field">
      <label>Subject Keywords</label>
      <textarea id="keywords" placeholder="urgent, asap, time-sensitive"></textarea>
      <div class="help">Comma-separated. Emails with these words in the subject will trigger an alert regardless of sender.</div>
    </div>

    <div class="field">
      <label>Exclude Domains (keyword alerts only)</label>
      <textarea id="excludeDomains" rows="2" placeholder="noreply, no-reply, mailer-daemon"></textarea>
      <div class="help">Comma-separated. Emails from these domains won't trigger keyword alerts (VIP sender alerts still apply).</div>
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
      document.getElementById('lookbackCount').value = s.lookbackCount || 50;
      document.getElementById('excludeDomains').value = s.excludeDomains || '';
    }).loadSettings();

    function save() {
      const settings = {
        checkFrequency: document.getElementById('checkFrequency').value,
        alertEmail: document.getElementById('alertEmail').value,
        keywords: document.getElementById('keywords').value,
        quietStart: document.getElementById('quietStart').value,
        quietEnd: document.getElementById('quietEnd').value,
        quietEnabled: document.getElementById('quietEnabled').checked,
        lookbackCount: parseInt(document.getElementById('lookbackCount').value, 10) || 50,
        excludeDomains: document.getElementById('excludeDomains').value,
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
  '<div class="header"><div class="icon">🕷</div>' +
  '<h2>VIP Priority Alert</h2><p>Quick Reference Guide</p></div>' +
  '<div class="content">' +
  '<div class="section"><h3>Quick Start</h3><ol>' +
  '<li>Open <strong>⚙️ Settings</strong> and add your VIP senders — top clients, your boss, key partners</li>' +
  '<li>Optionally add keywords to flag emails by subject regardless of sender</li>' +
  '<li>Click <strong>▶️ Start Monitoring</strong> to activate</li>' +
  '<li>VIP emails appear in the dashboard with sender, subject, and a direct Gmail link</li>' +
  '<li>Enable auto-schedule so new VIP messages are caught automatically</li>' +
  '</ol></div>' +
  '<div class="section"><h3>Settings Guide</h3>' +
  '<div class="setting"><strong>VIP Senders</strong><span>Email addresses to treat as high priority. Comma-separated. These people always get flagged.</span></div>' +
  '<div class="setting"><strong>Keywords</strong><span>Flag any email whose subject contains these words, regardless of who sent it. Comma-separated.</span></div>' +
  '<div class="setting"><strong>Check Frequency</strong><span>How often to scan for new VIP emails when auto-schedule is on</span></div>' +
  '<div class="setting"><strong>Email Notification</strong><span>Sends you an email alert when a new VIP message is detected</span></div>' +
  '</div>' +
  '<div class="section"><h3>Tips</h3>' +
  '<div class="tip">Combine <strong>senders + keywords</strong> for precision — e.g. only flag emails from your team that contain "urgent"</div>' +
  '<div class="tip">The dashboard keeps a full history — use it to review your response times to VIP contacts</div>' +
  '<div class="tip">Use <strong>Test Run</strong> to see what recent emails would have been flagged before activating</div>' +
  '</div></div>' +
  '<div class="footer">TAKScripts · takscripts.store</div>' +
  '</body></html>';
}
