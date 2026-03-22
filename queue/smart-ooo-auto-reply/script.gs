/**
 * Smart OOO Auto-Reply by TAKScripts
 * ====================================
 * Automatically replies to incoming emails when you're out of office.
 *
 * Features:
 * - Custom menu in Gmail/Sheets — no need to touch the script editor
 * - Beautiful settings sidebar for configuration
 * - Set OOO date range with start/end dates
 * - Custom reply message with dynamic {{name}}, {{start}}, {{end}} variables
 * - Only replies once per sender (prevents spam loops)
 * - Skips newsletters, promotions, and automated emails
 * - Logs all auto-replies to a Google Sheet
 * - Runs every 5 minutes via automatic trigger
 * - Idempotent — safe to run multiple times
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

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
    .addItem('⚙️ OOO Settings', 'showSettings')
    .addItem('▶️ Start OOO', 'startOOO')
    .addItem('⏹ Stop OOO', 'stopOOO')
    .addSeparator()
    .addItem('🧪 Test Run (no replies sent)', 'testRun')
    .addItem('📊 View Reply Log', 'viewLog')
    .addSeparator()
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the settings sidebar with branded UI.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(getSettingsHtml())
    .setTitle('OOO Auto-Reply Settings')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the about dialog.
 */
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

/**
 * Navigates to or creates the reply log sheet.
 */
function viewLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('OOO Reply Log');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No reply log yet. Start your OOO and replies will be logged here.');
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
  props.setProperty('ooo_settings', JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('ooo_settings');
  if (!raw) {
    return {
      startDate: '',
      endDate: '',
      subject: "Out of Office — I'll be back soon",
      message: "Hi {{name}},\n\nThanks for reaching out! I'm currently out of the office from {{start}} to {{end}} and won't be checking email regularly.\n\nI'll get back to you as soon as I return.\n\nBest,\n[Your Name]",
      skipDomains: 'noreply, no-reply, notifications, mailer-daemon, newsletter',
    };
  }
  return JSON.parse(raw);
}

// ═══════════════════════════════════════════
// CORE ENGINE
// ═══════════════════════════════════════════

/**
 * Main function — checks for new emails and sends OOO replies.
 */
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

    lastMessage.reply(replyBody, {
      subject: settings.subject,
    });

    replied.add(senderEmail);
    replyCount++;
    logReply(senderEmail, senderName, lastMessage.getSubject());
    Logger.log('Replied to: ' + senderEmail);
  }

  props.setProperty('ooo_replied', JSON.stringify([...replied]));
  Logger.log('Done. Sent ' + replyCount + ' auto-replies.');
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

/**
 * Logs reply to the "OOO Reply Log" sheet. Creates it if it doesn't exist.
 * Idempotent — safe to call multiple times.
 */
function logReply(email, name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('OOO Reply Log');

  if (!sheet) {
    sheet = ss.insertSheet('OOO Reply Log');
    const headers = ['Timestamp', 'Sender Email', 'Sender Name', 'Original Subject'];
    sheet.appendRow(headers);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
      .setBackground('#1a1a1a')
      .setFontColor('#C9A84C')
      .setFontFamily('Roboto Mono');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 300);
  }

  sheet.appendRow([new Date(), email, name, subject]);
}

// ═══════════════════════════════════════════
// CONTROLS
// ═══════════════════════════════════════════

/**
 * Activates the OOO auto-reply. Creates a 5-minute trigger.
 */
function startOOO() {
  const settings = loadSettings();
  if (!settings.startDate || !settings.endDate) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No dates configured.\n\nOpen TAKScripts → OOO Settings to set your dates first.'
    );
    return;
  }

  stopOOO();

  ScriptApp.newTrigger('checkAndReply')
    .timeBased()
    .everyMinutes(5)
    .create();

  checkAndReply();

  SpreadsheetApp.getUi().alert(
    '✅ OOO Auto-Reply is ACTIVE\n\n' +
    'Running from ' + settings.startDate + ' to ' + settings.endDate + '\n' +
    'Checking every 5 minutes for new emails.\n\n' +
    'To stop: TAKScripts → Stop OOO'
  );
}

/**
 * Deactivates the OOO auto-reply.
 */
function stopOOO() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'checkAndReply') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  PropertiesService.getScriptProperties().deleteProperty('ooo_replied');
}

/**
 * Dry run — shows what would happen without sending any replies.
 */
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
  SpreadsheetApp.getUi().alert(output);
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
    .field input, .field textarea {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      font-family: inherit;
      transition: border-color 0.2s;
    }
    .field input:focus, .field textarea:focus {
      outline: none;
      border-color: #C9A84C;
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);
    }
    .field textarea { min-height: 120px; resize: vertical; line-height: 1.5; }
    .field .help {
      font-size: 11px;
      color: #999;
      margin-top: 4px;
      line-height: 1.4;
    }
    .dates { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
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
      margin-top: 8px;
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

    <button class="btn btn-primary" onclick="save()">Save Settings</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>

    <div id="status" class="status"></div>
  </div>

  <script>
    // Load saved settings on open
    google.script.run.withSuccessHandler(function(settings) {
      document.getElementById('startDate').value = settings.startDate || '';
      document.getElementById('endDate').value = settings.endDate || '';
      document.getElementById('subject').value = settings.subject || '';
      document.getElementById('message').value = settings.message || '';
      document.getElementById('skipDomains').value = settings.skipDomains || '';
    }).loadSettings();

    function save() {
      var settings = {
        startDate: document.getElementById('startDate').value,
        endDate: document.getElementById('endDate').value,
        subject: document.getElementById('subject').value,
        message: document.getElementById('message').value,
        skipDomains: document.getElementById('skipDomains').value,
      };

      var statusEl = document.getElementById('status');
      statusEl.className = 'status';
      statusEl.style.display = 'none';

      google.script.run
        .withSuccessHandler(function() {
          statusEl.textContent = '✓ Settings saved successfully';
          statusEl.className = 'status success';
        })
        .withFailureHandler(function(err) {
          statusEl.textContent = '✕ Error: ' + err.message;
          statusEl.className = 'status error';
        })
        .saveSettings(settings);
    }
  </script>
</body>
</html>`;
}
