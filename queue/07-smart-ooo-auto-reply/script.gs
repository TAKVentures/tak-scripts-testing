/**
 * Smart OOO Auto-Reply by TAKScripts
 * ====================================
 * Automatically replies to incoming emails when you're out of office.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Dashboard stats: Total Replies · Unique Senders · This Trip · Today · Days Left
 * - Set OOO date range with start/end dates
 * - Custom reply message with dynamic {{name}}, {{start}}, {{end}} variables
 * - Only replies once per sender (prevents spam loops)
 * - Skips newsletters, promotions, and automated emails
 * - All replies logged to a branded dashboard sheet
 * - Runs every 5 minutes via automatic trigger
 * - Idempotent — safe to run multiple times
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const DASHBOARD_SHEET_NAME = '📊 OOO Dashboard';
const DASHBOARD_HEADER_ROW = 3;

const BRAND = {
  darkBg: '#1A1A1A',
  gold: '#C9A84C',
  white: '#FFFFFF',
  lightGray: '#F9F9F9',
  successBg: '#E8F5E9', successText: '#2E7D32',
  headerFont: 'Roboto Mono', bodyFont: 'Roboto',
};

const LOG_HEADERS = ['Timestamp', 'Sender Email', 'Sender Name', 'Original Subject'];


// ═══════════════════════════════════════════
// MENU & UI
// ═══════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🕷 TAKScripts')
    .addItem('⚙️ OOO Settings', 'showSettings')
    .addItem('▶️ Start OOO', 'startOOO')
    .addItem('⏹ Stop OOO', 'stopOOO')
    .addSeparator()
    .addItem('🧪 Test Run (no replies sent)', 'testRun')
    .addItem('📊 View Dashboard', 'viewDashboard')
    .addItem('🔄 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('OOO Auto-Reply Settings')
    .setWidth(360);
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
      <h2 style="margin: 0 0 4px; font-size: 18px;">Smart OOO Auto-Reply</h2>
      <p style="color: #666; font-size: 12px; margin: 0 0 16px;">Version 2.0 · by TAK Ventures</p>
      <hr style="border: none; border-top: 1px solid #eee; margin: 16px 0;">
      <p style="font-size: 13px; color: #333; line-height: 1.6;">
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
  `).setWidth(300).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'About TAKScripts');
}


// ═══════════════════════════════════════════
// SETTINGS MANAGEMENT
// ═══════════════════════════════════════════

function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ooo_settings', JSON.stringify(settings));
  return { success: true };
}

function getDefaultSettings_() {
  return {
    startDate: '',
    endDate: '',
    subject: "Out of Office — I'll be back soon",
    message: "Hi {{name}},\n\nThanks for reaching out! I'm currently out of the office from {{start}} to {{end}} and won't be checking email regularly.\n\nI'll get back to you as soon as I return.\n\nBest,\n[Your Name]",
    skipDomains: 'noreply, no-reply, notifications, mailer-daemon, newsletter',
  };
}

function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('ooo_settings');
  if (!raw) return getDefaultSettings_();
  const saved = JSON.parse(raw);
  const defaults = getDefaultSettings_();
  for (const key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}


// ═══════════════════════════════════════════
// DASHBOARD MANAGEMENT
// ═══════════════════════════════════════════

/**
 * Gets or creates the OOO Dashboard sheet with stats bar and log headers.
 * Idempotent — safe to call multiple times.
 */
function getOrCreateDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    // Migration: rename old log sheet if it exists
    const oldSheet = ss.getSheetByName('OOO Reply Log');
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
    .setValues([['—', '—', '—', '—']])
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
    .setValues([['TOTAL REPLIES', 'UNIQUE SENDERS', 'THIS TRIP', 'TODAY']])
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
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 250); // Sender Email
  sheet.setColumnWidth(3, 150); // Sender Name
  sheet.setColumnWidth(4, 300); // Original Subject

  return sheet;
}

/**
 * Recalculates all stats from log data and settings, updates the stats bar.
 */
function refreshDashboardStats() {
  const sheet = getOrCreateDashboard_();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const lastRow = sheet.getLastRow();

  let totalReplies = 0;
  const senders = new Set();
  let thisTrip = 0;
  let today = 0;

  const settings = loadSettings();
  const tripStart = settings.startDate ? new Date(settings.startDate + 'T00:00:00') : null;
  const tripEnd = settings.endDate ? new Date(settings.endDate + 'T23:59:59') : null;
  const now = new Date();

  if (lastRow >= dataStartRow) {
    const numRows = lastRow - dataStartRow + 1;
    const data = sheet.getRange(dataStartRow, 1, numRows, LOG_HEADERS.length).getValues();

    for (const row of data) {
      if (!row[0]) continue; // Skip empty rows
      totalReplies++;
      const email = (row[1] || '').toString().toLowerCase();
      if (email) senders.add(email);

      const ts = row[0];
      if (ts instanceof Date) {
        // This trip
        if (tripStart && tripEnd && ts >= tripStart && ts <= tripEnd) thisTrip++;

        // Today
        if (ts.getDate() === now.getDate() &&
            ts.getMonth() === now.getMonth() &&
            ts.getFullYear() === now.getFullYear()) {
          today++;
        }
      }
    }
  }

  sheet.getRange(1, 1, 1, 4).setValues([[
    totalReplies || '—',
    senders.size || '—',
    thisTrip || '—',
    today || '—',
  ]]);
}


// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

function checkAndReply() {
  const settings = loadSettings();
  if (!settings.startDate || !settings.endDate) {
    Logger.log('No OOO dates configured. Open Settings from the TAKScripts menu.');
    return;
  }

  const now = new Date();
  const start = new Date(settings.startDate + 'T00:00:00');
  const end = new Date(settings.endDate + 'T23:59:59');

  if (now < start || now > end) {
    Logger.log('Outside OOO date range. Skipping.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const repliedRaw = props.getProperty('ooo_replied') || '[]';
  const replied = new Set(JSON.parse(repliedRaw));

  const skipDomains = (settings.skipDomains || '')
    .split(',')
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  const threads = GmailApp.search('is:unread is:inbox newer_than:1d');
  let replyCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const senderEmail = extractEmail(lastMessage.getFrom());
    const senderName = extractName(lastMessage.getFrom());

    if (replied.has(senderEmail)) continue;
    if (shouldSkip(senderEmail, lastMessage, skipDomains)) continue;

    const myEmail = Session.getActiveUser().getEmail();
    if (senderEmail === myEmail) continue;

    const replyBody = settings.message
      .replace(/\{\{name\}\}/g, senderName || 'there')
      .replace(/\{\{start\}\}/g, formatDate(start))
      .replace(/\{\{end\}\}/g, formatDate(end));

    lastMessage.reply(replyBody, { subject: settings.subject });

    replied.add(senderEmail);
    replyCount++;
    logReply(senderEmail, senderName, lastMessage.getSubject());
    Logger.log('Replied to: ' + senderEmail);
  }

  props.setProperty('ooo_replied', JSON.stringify([...replied]));
  Logger.log('Done. Sent ' + replyCount + ' auto-replies.');
}


// ═══════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════

/**
 * Logs a reply to the OOO Dashboard and refreshes stats.
 */
function logReply(email, name, subject) {
  const sheet = getOrCreateDashboard_();
  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const newRow = Math.max(sheet.getLastRow() + 1, dataStartRow);

  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setValues([
    [new Date(), email, name || '', subject || ''],
  ]);

  // Row styling
  const isEven = (newRow - DASHBOARD_HEADER_ROW) % 2 === 0;
  sheet.getRange(newRow, 1, 1, LOG_HEADERS.length)
    .setFontFamily(BRAND.bodyFont)
    .setFontSize(10)
    .setBackground(isEven ? BRAND.lightGray : BRAND.white);
  sheet.setRowHeight(newRow, 30);

  sheet.getRange(newRow, 1).setNumberFormat('MMM d, yyyy h:mm a');

  // Highlight the action indicator on sender email
  sheet.getRange(newRow, 2)
    .setBackground(BRAND.successBg)
    .setFontColor(BRAND.successText);

  refreshDashboardStats();
}


// ═══════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════

function extractEmail(fromField) {
  const match = fromField.match(/<(.+?)>/);
  return (match ? match[1] : fromField).toLowerCase().trim();
}

function extractName(fromField) {
  const match = fromField.match(/^"?([^"<]+)"?\s*</);
  return match ? match[1].trim().split(' ')[0] : '';
}

function shouldSkip(email, message, skipDomains) {
  for (const skip of skipDomains) {
    if (email.includes(skip)) return true;
  }

  const labels = message.getThread().getLabels();
  const skipCategories = ['promotions', 'social', 'updates', 'forums'];
  for (const label of labels) {
    if (skipCategories.includes(label.getName().toLowerCase())) return true;
  }

  if (message.getHeader('List-Unsubscribe')) return true;
  return false;
}

function formatDate(date) {
  const months = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}


// ═══════════════════════════════════════════
// CONTROLS
// ═══════════════════════════════════════════

function startOOO() {
  const settings = loadSettings();
  if (!settings.startDate || !settings.endDate) {
    try {
      SpreadsheetApp.getUi().alert(
        '⚠️ No dates configured.\n\nOpen TAKScripts → OOO Settings to set your dates first.'
      );
    } catch(e) {
      Logger.log('⚠️ No dates configured.\n\nOpen TAKScripts → OOO Settings to set your dates first.');
    }
    return;
  }

  stopOOO();
  getOrCreateDashboard_();

  ScriptApp.newTrigger('checkAndReply')
    .timeBased()
    .everyMinutes(5)
    .create();

  checkAndReply();

  try {
    SpreadsheetApp.getUi().alert(
      '✅ OOO Auto-Reply is ACTIVE\n\n' +
      'Running from ' + settings.startDate + ' to ' + settings.endDate + '\n' +
      'Checking every 5 minutes for new emails.\n\n' +
      'To stop: TAKScripts → Stop OOO'
    );
  } catch(e) {
    Logger.log('✅ OOO Auto-Reply is ACTIVE\n\nRunning from ' + settings.startDate + ' to ' + settings.endDate + '\nChecking every 5 minutes for new emails.\n\nTo stop: TAKScripts → Stop OOO');
  }
}

function stopOOO() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'checkAndReply') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  PropertiesService.getScriptProperties().deleteProperty('ooo_replied');
}

function testRun() {
  const threads = GmailApp.search('is:unread is:inbox newer_than:1d');
  const results = [];

  results.push('🧪 TEST RUN — No replies will be sent');
  results.push('Found ' + threads.length + ' unread threads from the last 24 hours.');
  results.push('---');

  const settings = loadSettings();
  const skipDomains = (settings.skipDomains || '')
    .split(',')
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const senderEmail = extractEmail(lastMessage.getFrom());
    const senderName = extractName(lastMessage.getFrom());
    const skip = shouldSkip(senderEmail, lastMessage, skipDomains);

    results.push('From: ' + (senderName || '(no name)') + ' <' + senderEmail + '>');
    results.push('Subject: ' + lastMessage.getSubject());
    results.push(skip ? '→ SKIP (filtered)' : '→ WOULD REPLY with OOO message');
    results.push('');
  }

  if (threads.length === 0) {
    results.push('No unread emails to process. All clear!');
  }

  const output = results.join('\n');
  Logger.log(output);
  try {
    SpreadsheetApp.getUi().alert(output);
  } catch(e) {
    Logger.log(output);
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
    body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: #fafafa; color: #1a1a1a; font-size: 13px; }
    .header { background: #1a1a1a; color: white; padding: 20px 16px; text-align: center; }
    .header .logo { font-size: 24px; margin-bottom: 4px; }
    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }
    .header .brand { color: #C9A84C; }
    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }
    .form { padding: 16px; }
    .field { margin-bottom: 16px; }
    .field label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 6px; }
    .field input, .field textarea {
      width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px;
      font-size: 13px; font-family: inherit; transition: border-color 0.2s;
    }
    .field input:focus, .field textarea:focus {
      outline: none; border-color: #C9A84C; box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea { min-height: 120px; resize: vertical; line-height: 1.5; }
    .field .help { font-size: 11px; color: #999; margin-top: 4px; line-height: 1.4; }
    .dates { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .btn { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; letter-spacing: 0.5px; }
    .btn-primary { background: #1a1a1a; color: #C9A84C; border: 1px solid #C9A84C; }
    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }
    .btn-secondary { background: white; color: #666; border: 1px solid #ddd; margin-top: 8px; }
    .btn-secondary:hover { border-color: #999; color: #333; }
    .status { text-align: center; padding: 8px; font-size: 12px; margin-top: 8px; border-radius: 6px; display: none; }
    .status.success { display: block; background: #e8f5e9; color: #2e7d32; }
    .status.error { display: block; background: #ffebee; color: #c62828; }
    .divider { border-top: 1px solid #eee; margin: 20px 0; }
    .variables { background: #f5f5f5; border-radius: 6px; padding: 10px 12px; margin-top: 8px; }
    .variables code { display: inline-block; background: #e8e8e8; padding: 1px 6px; border-radius: 3px; font-size: 11px; margin: 2px; }
  </style>
</head>
<body>
  <div class="header">
    <div class="logo">🕷</div>
    <h1><span class="brand">TAK</span>Scripts</h1>
    <div class="sub">Smart OOO Auto-Reply · Settings</div>
  </div>

  <div class="form">
    <div class="field">
      <label>OOO Dates</label>
      <div class="dates">
        <div>
          <input type="date" id="startDate" />
          <div class="help">Start</div>
        </div>
        <div>
          <input type="date" id="endDate" />
          <div class="help">End</div>
        </div>
      </div>
    </div>

    <div class="field">
      <label>Reply Subject</label>
      <input type="text" id="subject" placeholder="Out of Office — I'll be back soon" />
    </div>

    <div class="field">
      <label>Reply Message</label>
      <textarea id="message" placeholder="Hi {{name}}, thanks for reaching out..."></textarea>
      <div class="variables">
        <div class="help" style="margin-bottom: 4px;">Template variables:</div>
        <code>{{name}}</code> sender's first name
        <code>{{start}}</code> your start date
        <code>{{end}}</code> your end date
      </div>
    </div>

    <div class="field">
      <label>Skip These Senders</label>
      <input type="text" id="skipDomains" placeholder="noreply, newsletter, mailer-daemon" />
      <div class="help">Comma-separated. Emails containing these words will be skipped.</div>
    </div>

    <div class="divider"></div>

    <button id="saveBtn" class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    google.script.run.withSuccessHandler(function(settings) {
      document.getElementById('startDate').value = settings.startDate || '';
      document.getElementById('endDate').value = settings.endDate || '';
      document.getElementById('subject').value = settings.subject || '';
      document.getElementById('message').value = settings.message || '';
      document.getElementById('skipDomains').value = settings.skipDomains || '';
    }).loadSettings();

    function save() {
      const settings = {
        startDate: document.getElementById('startDate').value,
        endDate: document.getElementById('endDate').value,
        subject: document.getElementById('subject').value,
        message: document.getElementById('message').value,
        skipDomains: document.getElementById('skipDomains').value,
      };

      const statusEl = document.getElementById('status');
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
