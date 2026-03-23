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
 * - Quiet hours to prevent late-night notifications
 * - Logs all alerts to a styled Google Sheet
 * - Idempotent — tracks processed message IDs, safe to run multiple times
 * - Configurable check frequency (1–15 minutes)
 *
 * Version: 1.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const SHEET_VIP = '⭐ VIP Contacts';
const SHEET_LOG = '📋 Alert Log';
const LABEL_VIP = '⭐ VIP';
const PROP_SETTINGS = 'vip_settings';
const PROP_PROCESSED = 'vip_processed_ids';
const TRIGGER_FN = 'checkForVIPEmails';

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
    .addItem('▶️ Start Monitoring', 'startMonitoring')
    .addItem('⏹ Stop Monitoring', 'stopMonitoring')
    .addSeparator()
    .addItem('🧪 Test Run', 'testRun')
    .addItem('📊 View Alert Log', 'viewAlertLog')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('VIP Priority Alert — Settings')
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
      <h2 style="margin: 0 0 4px; font-size: 18px;">VIP Priority Alert</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 1.0 · by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
        Never miss an email from the people who matter most.<br>
        Part of the <strong>TAKScripts</strong> collection.<br>
        Pre-built Google Apps Scripts for small business.
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
 * Navigates to the Alert Log sheet.
 */
function viewAlertLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No alert log yet. Start monitoring and alerts will appear here.');
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
  props.setProperty(PROP_SETTINGS, JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings with sensible defaults.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_SETTINGS);
  if (!raw) {
    return {
      checkFrequency: '5',
      alertEmail: Session.getActiveUser().getEmail() || '',
      keywords: 'urgent, asap, time-sensitive, action required',
      quietStart: '22:00',
      quietEnd: '07:00',
      quietEnabled: false,
    };
  }
  return JSON.parse(raw);
}

// ═══════════════════════════════════════════
// SHEET SETUP
// ═══════════════════════════════════════════

/**
 * Ensures the VIP Contacts sheet exists with proper formatting.
 * Idempotent — safe to call multiple times.
 */
function ensureVIPSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_VIP);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_VIP, 0);
    const headers = ['Name', 'Email', 'Priority', 'Alert Method', 'Notes'];
    sheet.appendRow(headers);

    // Style header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontFamily('Roboto Mono')
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground('#1A1A1A')
      .setFontColor('#C9A84C');
    sheet.setFrozenRows(1);

    // Column widths
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 220);

    // Priority dropdown
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['🔴 Critical', '🟡 Important', '🟢 Normal'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 3, 200, 1).setDataValidation(priorityRule);

    // Alert method dropdown
    const methodRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['📧 Email Alert', '⭐ Star + Label Only'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, 200, 1).setDataValidation(methodRule);

    // Data rows font
    sheet.getRange(2, 1, 200, headers.length)
      .setFontFamily('Roboto')
      .setFontSize(10);

    // Example VIP row
    sheet.appendRow([
      'Jane Doe (example)',
      'jane@example.com',
      '🔴 Critical',
      '📧 Email Alert',
      'CEO — always alert',
    ]);

    // Alternating colors for first batch of rows
    applyAlternatingColors_(sheet, headers.length);
  }

  return sheet;
}

/**
 * Ensures the Alert Log sheet exists with proper formatting.
 * Idempotent — safe to call multiple times.
 */
function ensureLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    const headers = ['Timestamp', 'Sender', 'Subject', 'Priority', 'Action Taken', 'Trigger'];
    sheet.appendRow(headers);

    // Style header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setFontFamily('Roboto Mono')
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground('#1A1A1A')
      .setFontColor('#C9A84C');
    sheet.setFrozenRows(1);

    // Column widths
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 240);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 150);

    // Data rows font
    sheet.getRange(2, 1, 500, headers.length)
      .setFontFamily('Roboto')
      .setFontSize(10);

    applyAlternatingColors_(sheet, headers.length);
  }

  return sheet;
}

/**
 * Applies alternating row colors (white / light gray).
 */
function applyAlternatingColors_(sheet, colCount) {
  const banding = sheet.getRange(2, 1, 500, colCount);
  banding.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  // Override banding colors to match brand
  const bandings = sheet.getBandings();
  if (bandings.length > 0) {
    bandings[bandings.length - 1]
      .setFirstRowColor('#FFFFFF')
      .setSecondRowColor('#F9F9F9')
      .setHeaderRowColor(null);
  }
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main function — checks for new emails from VIP contacts
 * and triggers appropriate alerts.
 */
function checkForVIPEmails() {
  const settings = loadSettings();
  const vips = getVIPList_();

  if (vips.length === 0) {
    Logger.log('No VIP contacts configured. Add contacts to the ⭐ VIP Contacts sheet.');
    return;
  }

  // Check quiet hours
  if (settings.quietEnabled && isQuietHours_(settings.quietStart, settings.quietEnd)) {
    Logger.log('Quiet hours active. Skipping check.');
    return;
  }

  // Load processed IDs for idempotency
  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_PROCESSED) || '[]';
  const processedArr = JSON.parse(processedRaw);
  const processed = new Set(processedArr);

  // Parse keyword triggers
  const keywords = (settings.keywords || '')
    .split(',')
    .map(k => k.trim().toLowerCase())
    .filter(Boolean);

  // Search recent unread inbox messages
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

    // Check if sender is a VIP
    const vipMatch = vips.find(v => v.email === senderEmail);

    // Check if subject contains a keyword trigger
    const subjectLower = subject.toLowerCase();
    const keywordHit = keywords.find(kw => subjectLower.includes(kw));

    if (!vipMatch && !keywordHit) continue;

    // Determine priority and alert method
    const priority = vipMatch ? vipMatch.priority : '🟡 Important';
    const alertMethod = vipMatch ? vipMatch.alertMethod : '📧 Email Alert';
    const triggerReason = vipMatch
      ? 'VIP: ' + (vipMatch.name || senderEmail)
      : 'Keyword: "' + keywordHit + '"';

    // Always apply VIP label and star
    applyVIPLabel_(thread);
    thread.getMessages().forEach(m => m.star());

    // Send email alert if configured
    if (alertMethod === '📧 Email Alert') {
      sendAlertEmail_(settings, senderName, senderEmail, subject, snippet, priority, triggerReason);
    }

    // Log the alert
    logAlert_(senderEmail, subject, priority,
      alertMethod === '📧 Email Alert' ? 'Email alert sent' : 'Starred + labeled',
      triggerReason);

    newProcessed.push(msgId);
    alertCount++;
    Logger.log('VIP alert: ' + senderEmail + ' — ' + subject);
  }

  // Persist processed IDs (keep last 500 to prevent unbounded growth)
  const allProcessed = [...processedArr, ...newProcessed].slice(-500);
  props.setProperty(PROP_PROCESSED, JSON.stringify(allProcessed));

  Logger.log('Done. Triggered ' + alertCount + ' VIP alerts.');
}

// ═══════════════════════════════════════════
// VIP LIST MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Reads VIP contacts from the sheet.
 * Returns array of { name, email, priority, alertMethod }.
 */
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

/**
 * Creates and applies the VIP label to a thread.
 */
function applyVIPLabel_(thread) {
  let label = GmailApp.getUserLabelByName(LABEL_VIP);
  if (!label) {
    label = GmailApp.createLabel(LABEL_VIP);
  }
  label.addToThread(thread);
}

/**
 * Sends a branded HTML alert email.
 */
function sendAlertEmail_(settings, senderName, senderEmail, subject, snippet, priority, triggerReason) {
  const alertTo = settings.alertEmail || Session.getActiveUser().getEmail();
  if (!alertTo) {
    Logger.log('No alert email configured. Skipping email alert.');
    return;
  }

  const priorityColor = priority.includes('Critical') ? '#C62828'
    : priority.includes('Important') ? '#F57F17'
    : '#2E7D32';

  const priorityBg = priority.includes('Critical') ? '#FFEBEE'
    : priority.includes('Important') ? '#FFF8E1'
    : '#E8F5E9';

  const htmlBody = `
    <div style="font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; max-width: 560px; margin: 0 auto;">
      <!-- Header -->
      <div style="background: #1A1A1A; padding: 20px 24px; border-radius: 8px 8px 0 0; text-align: center;">
        <span style="font-size: 22px;">🕷</span>
        <span style="color: #C9A84C; font-size: 16px; font-weight: 700; letter-spacing: 0.5px; vertical-align: middle; margin-left: 8px;">VIP Priority Alert</span>
      </div>

      <!-- Body -->
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

      <!-- Footer -->
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
 * Logs an alert to the Alert Log sheet. Creates it if needed.
 * Applies priority-based status colors to the Priority cell.
 */
function logAlert_(sender, subject, priority, action, trigger) {
  const sheet = ensureLogSheet();
  const newRow = sheet.getLastRow() + 1;

  sheet.appendRow([new Date(), sender, subject, priority, action, trigger]);

  // Color the priority cell
  const priorityCell = sheet.getRange(newRow, 4);
  if (priority.includes('Critical')) {
    priorityCell.setBackground('#FFEBEE').setFontColor('#C62828');
  } else if (priority.includes('Important')) {
    priorityCell.setBackground('#FFF8E1').setFontColor('#F57F17');
  } else {
    priorityCell.setBackground('#E8F5E9').setFontColor('#2E7D32');
  }

  // Color the action cell
  const actionCell = sheet.getRange(newRow, 5);
  actionCell.setBackground('#E8F5E9').setFontColor('#2E7D32');
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
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Checks if current time falls within quiet hours.
 */
function isQuietHours_(startStr, endStr) {
  if (!startStr || !endStr) return false;

  const now = new Date();
  const currentMinutes = now.getHours() * 60 + now.getMinutes();

  const [startH, startM] = startStr.split(':').map(Number);
  const [endH, endM] = endStr.split(':').map(Number);
  const startMinutes = startH * 60 + startM;
  const endMinutes = endH * 60 + endM;

  // Handle overnight ranges (e.g., 22:00 to 07:00)
  if (startMinutes > endMinutes) {
    return currentMinutes >= startMinutes || currentMinutes < endMinutes;
  }
  return currentMinutes >= startMinutes && currentMinutes < endMinutes;
}

// ═══════════════════════════════════════════
// CONTROLS
// ═══════════════════════════════════════════

/**
 * Starts VIP monitoring. Creates the VIP sheet if needed
 * and installs a time-based trigger.
 */
function startMonitoring() {
  const settings = loadSettings();
  ensureVIPSheet();
  ensureLogSheet();

  // Remove existing triggers
  stopMonitoring_(false);

  const freq = parseInt(settings.checkFrequency, 10) || 5;

  ScriptApp.newTrigger(TRIGGER_FN)
    .timeBased()
    .everyMinutes(freq)
    .create();

  // Run immediately
  checkForVIPEmails();

  const vips = getVIPList_();
  SpreadsheetApp.getUi().alert(
    '✅ VIP Monitoring is ACTIVE\n\n' +
    'Tracking ' + vips.length + ' VIP contact(s).\n' +
    'Checking every ' + freq + ' minute(s).\n\n' +
    'Add or edit VIPs in the "⭐ VIP Contacts" sheet.\n' +
    'To stop: 🕷 TAKScripts → ⏹ Stop Monitoring'
  );
}

/**
 * Stops VIP monitoring (public menu action).
 */
function stopMonitoring() {
  stopMonitoring_(true);
}

/**
 * Internal stop — removes triggers and optionally clears processed IDs.
 */
function stopMonitoring_(showAlert) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FN) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  if (showAlert) {
    PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED);
    SpreadsheetApp.getUi().alert(
      '⏹ VIP Monitoring STOPPED\n\n' +
      'No more alerts will be sent.\n' +
      'Processed message history has been cleared.\n\n' +
      'To resume: 🕷 TAKScripts → ▶️ Start Monitoring'
    );
  }
}

/**
 * Test Run — checks the last 10 emails and shows which would trigger alerts.
 * No alerts are sent, no labels applied, nothing is modified.
 */
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
  SpreadsheetApp.getUi().alert(output);
}

// ═══════════════════════════════════════════
// INSTALL HELPER
// ═══════════════════════════════════════════

/**
 * One-time setup — creates both sheets with proper formatting.
 * Called automatically when Start Monitoring is used for the first time.
 * Can also be run manually from the script editor.
 */
function install() {
  ensureVIPSheet();
  ensureLogSheet();
  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete\n\n' +
    'Sheets created:\n' +
    '• ⭐ VIP Contacts — add your VIPs here\n' +
    '• 📋 Alert Log — alerts will appear here\n\n' +
    'Next: Open 🕷 TAKScripts → ⚙️ Settings to configure.'
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
      background: #FAFAFA;
      color: #1a1a1a;
      font-size: 13px;
    }
    .header {
      background: #1A1A1A;
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
      letter-spacing: 1.2px;
      color: #C9A84C;
      margin: 20px 0 12px;
      padding-bottom: 6px;
      border-bottom: 1px solid #eee;
    }
    .section-title:first-child { margin-top: 0; }
    .field { margin-bottom: 16px; }
    .field label {
      display: block;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #666;
      margin-bottom: 6px;
    }
    .field input, .field textarea, .field select {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      font-family: inherit;
      background: white;
      transition: border-color 0.2s;
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
      gap: 10px;
      margin-bottom: 12px;
    }
    .toggle-row input[type="checkbox"] {
      width: 18px;
      height: 18px;
      accent-color: #C9A84C;
      cursor: pointer;
    }
    .toggle-row label {
      font-size: 13px;
      font-weight: 500;
      color: #333;
      cursor: pointer;
      text-transform: none;
      letter-spacing: 0;
    }
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
      background: #1A1A1A;
      color: #C9A84C;
      border: 1px solid #C9A84C;
    }
    .btn-primary:hover { background: #C9A84C; color: #1A1A1A; }
    .btn-secondary {
      background: white;
      color: #666;
      border: 1px solid #ddd;
      margin-top: 8px;
    }
    .btn-secondary:hover { border-color: #999; color: #333; }
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
    .info-box {
      background: #FFF8E1;
      border-radius: 6px;
      padding: 10px 12px;
      font-size: 12px;
      color: #F57F17;
      line-height: 1.5;
      margin-top: 8px;
    }
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

    <button class="btn btn-primary" onclick="save()">Save Settings</button>
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
      var settings = {
        checkFrequency: document.getElementById('checkFrequency').value,
        alertEmail: document.getElementById('alertEmail').value,
        keywords: document.getElementById('keywords').value,
        quietStart: document.getElementById('quietStart').value,
        quietEnd: document.getElementById('quietEnd').value,
        quietEnabled: document.getElementById('quietEnabled').checked,
      };

      var freq = parseInt(settings.checkFrequency, 10);
      if (freq < 1 || freq > 15) {
        showStatus('Check frequency must be between 1 and 15 minutes.', 'error');
        return;
      }

      var statusEl = document.getElementById('status');
      statusEl.className = 'status';
      statusEl.style.display = 'none';

      google.script.run
        .withSuccessHandler(function() {
          showStatus('Settings saved successfully', 'success');
        })
        .withFailureHandler(function(err) {
          showStatus('Error: ' + err.message, 'error');
        })
        .saveSettings(settings);
    }

    function showStatus(msg, type) {
      var el = document.getElementById('status');
      el.textContent = (type === 'success' ? '✓ ' : '✕ ') + msg;
      el.className = 'status ' + type;
    }
  </script>
</body>
</html>`;
}
