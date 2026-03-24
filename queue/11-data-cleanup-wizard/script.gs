/**
 * Data Cleanup Wizard by TAKScripts
 * ====================================
 * One-click data cleanup toolkit for any Google Sheet.
 *
 * Features:
 * - Custom menu in Sheets — no need to touch the script editor
 * - Beautiful sidebar with cleanup options and scope selector
 * - Remove duplicate rows (configurable columns)
 * - Trim whitespace (leading, trailing, extra internal spaces)
 * - Fix capitalization (Title Case, UPPER, lower, Sentence case)
 * - Standardize dates to a consistent format
 * - Remove empty rows and columns
 * - Standardize phone numbers
 * - Fix email formatting (lowercase, trim)
 * - Remove special characters
 * - Preview mode — highlights changes in yellow without modifying data
 * - Automatic backup sheet before any changes
 * - Undo via backup sheet
 * - Results dialog with full stats
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */

// ═══════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════

const DASHBOARD_SHEET_NAME = '\uD83D\uDCCA Cleanup Dashboard';
const DASHBOARD_HEADER_ROW = 3;

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
    .addItem('🧹 Run All Cleanups', 'runAllCleanups')
    .addItem('🧪 Preview Changes', 'previewChanges')
    .addSeparator()
    .addItem('🔁 Remove Duplicates', 'menuRemoveDuplicates')
    .addItem('✂️ Trim Whitespace', 'menuTrimWhitespace')
    .addItem('🔤 Fix Capitalization', 'menuFixCapitalization')
    .addItem('📅 Standardize Dates', 'menuStandardizeDates')
    .addItem('🗑 Remove Empty Rows & Columns', 'menuRemoveEmpties')
    .addItem('📞 Standardize Phone Numbers', 'menuStandardizePhones')
    .addItem('📧 Fix Email Formatting', 'menuFixEmails')
    .addItem('🚫 Remove Special Characters', 'menuRemoveSpecialChars')
    .addSeparator()
    .addItem('📊 View Dashboard', 'viewDashboard')
    .addItem('🔄 Refresh Stats', 'refreshDashboardStats')
    .addSeparator()
    .addItem('⚙️ Cleanup Settings', 'showSidebar')
    .addItem('↩️ Undo (Restore Backup)', 'undoFromBackup')
    .addSeparator()
    .addItem('❓ How to Use', 'showHelp')
    .addItem('ℹ️ About TAKScripts', 'showAbout')
    .addToUi();
}

/**
 * Shows the cleanup settings sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('Data Cleanup Wizard')
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
      <h2 style="margin: 0 0 4px; font-size: 18px;">Data Cleanup Wizard</h2>
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
 * Navigates to the Cleanup Dashboard sheet.
 */
function viewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    try {
      SpreadsheetApp.getUi().alert(
        'No cleanup history yet.\n\nRun a cleanup to generate dashboard entries.'
      );
    } catch(e) {
      Logger.log('No cleanup history yet.\n\nRun a cleanup to generate dashboard entries.');
    }
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
  props.setProperty('cleanup_settings', JSON.stringify(settings));
  return { success: true };
}

/**
 * Load saved settings.
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('cleanup_settings');
  if (!raw) return getDefaultSettings_();
  const saved = JSON.parse(raw);
  const defaults = getDefaultSettings_();
  for (const key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}

function getDefaultSettings_() {
  return {
    scope: 'sheet',
    dateFormat: 'MM/DD/YYYY',
    phoneFormat: '(XXX) XXX-XXXX',
    capStyle: 'title',
    specialCharsKeep: '.,!?@#$%&()-',
    dupColumns: '',
  };
}

// ═══════════════════════════════════════════
// BACKUP & UNDO
// ═══════════════════════════════════════════

/**
 * Creates a backup of the active sheet before making changes.
 * Returns the backup sheet name.
 */
function createBackup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const backupName = '💾 Backup ' + timestamp;

  const backup = sheet.copyTo(ss);
  backup.setName(backupName);

  // Move backup to end
  ss.setActiveSheet(backup);
  ss.moveActiveSheet(ss.getNumSheets());
  ss.setActiveSheet(sheet);

  return backupName;
}

/**
 * Restores data from the most recent backup sheet.
 */
function undoFromBackup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Find the most recent backup
  let latestBackup = null;
  for (const s of sheets) {
    if (s.getName().startsWith('💾 Backup')) {
      latestBackup = s;
    }
  }

  if (!latestBackup) {
    try {
      ui.alert('No backup found. Run a cleanup first to create a backup.');
    } catch(e) {
      Logger.log('No backup found. Run a cleanup first to create a backup.');
    }
    return;
  }

  let response;
  try {
    response = ui.alert(
      '↩️ Restore Backup',
      'This will replace the active sheet\'s data with the backup:\n\n"' + latestBackup.getName() + '"\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
  } catch(e) {
    Logger.log('↩️ Restore Backup\nThis will replace the active sheet\'s data with the backup:\n\n"' + latestBackup.getName() + '"\n\nContinue?');
  }

  if (response !== ui.Button.YES) return;

  const activeSheet = ss.getActiveSheet();
  const backupData = latestBackup.getDataRange().getValues();

  activeSheet.clearContents();
  activeSheet.clearFormats();

  if (backupData.length > 0 && backupData[0].length > 0) {
    activeSheet.getRange(1, 1, backupData.length, backupData[0].length).setValues(backupData);
  }

  // Copy formats from backup
  const backupFormats = latestBackup.getDataRange().getBackgrounds();
  const backupFontColors = latestBackup.getDataRange().getFontColors();
  if (backupData.length > 0 && backupData[0].length > 0) {
    activeSheet.getRange(1, 1, backupData.length, backupData[0].length).setBackgrounds(backupFormats);
    activeSheet.getRange(1, 1, backupData.length, backupData[0].length).setFontColors(backupFontColors);
  }

  try {
    ui.alert('✅ Backup restored successfully.\n\nYou can delete the backup sheet "' + latestBackup.getName() + '" if you no longer need it.');
  } catch(e) {
    Logger.log('✅ Backup restored successfully.\n\nYou can delete the backup sheet "' + latestBackup.getName() + '" if you no longer need it.');
  }
}

// ═══════════════════════════════════════════
// SCOPE HELPER
// ═══════════════════════════════════════════

/**
 * Returns the data range based on the configured scope.
 * @param {Object} settings — the loaded settings
 * @returns {GoogleAppsScript.Spreadsheet.Range}
 */
function getScopeRange_(settings) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const scope = settings.scope || 'sheet';

  if (scope === 'selection') {
    return sheet.getActiveRange();
  }

  if (scope === 'columns' && settings.scopeColumns) {
    const cols = settings.scopeColumns.split(',').map(c => c.trim().toUpperCase());
    if (cols.length === 0) return sheet.getDataRange();

    const firstCol = columnLetterToIndex_(cols[0]);
    const lastCol = columnLetterToIndex_(cols[cols.length - 1]);
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return sheet.getRange('A1');
    return sheet.getRange(1, firstCol, lastRow, lastCol - firstCol + 1);
  }

  // Default: entire sheet
  return sheet.getDataRange();
}

/**
 * Converts a column letter (A, B, AA) to a 1-based index.
 */
function columnLetterToIndex_(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

// ═══════════════════════════════════════════
// CORE CLEANUP FUNCTIONS
// ═══════════════════════════════════════════

/**
 * Removes duplicate rows. Returns count of duplicates removed.
 * @param {boolean} preview — if true, highlights duplicates instead of removing
 */
function removeDuplicates_(preview) {
  const settings = loadSettings();
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length <= 1) return 0; // Only header or empty

  // Determine which columns to compare
  let compareCols = [];
  if (settings.dupColumns && settings.dupColumns.trim()) {
    compareCols = settings.dupColumns.split(',').map(c => {
      const trimmed = c.trim().toUpperCase();
      return columnLetterToIndex_(trimmed) - 1; // 0-based
    });
  }

  const seen = new Set();
  const dupRows = []; // 0-based indices of duplicate rows

  for (let i = 1; i < data.length; i++) { // Skip header
    let key;
    if (compareCols.length > 0) {
      key = compareCols.map(c => String(data[i][c] || '').trim().toLowerCase()).join('|||');
    } else {
      key = data[i].map(v => String(v || '').trim().toLowerCase()).join('|||');
    }

    if (seen.has(key)) {
      dupRows.push(i);
    } else {
      seen.add(key);
    }
  }

  if (preview) {
    // Highlight duplicate rows in yellow
    for (const rowIdx of dupRows) {
      sheet.getRange(rowIdx + 1, 1, 1, data[0].length)
        .setBackground('#FFF9C4');
    }
  } else {
    // Remove from bottom to top to preserve indices
    for (let i = dupRows.length - 1; i >= 0; i--) {
      sheet.deleteRow(dupRows[i] + 1);
    }
  }

  return dupRows.length;
}

/**
 * Trims whitespace from all cells in scope. Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function trimWhitespace_(preview) {
  const settings = loadSettings();
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === 'string') {
        const original = data[r][c];
        const trimmed = original.trim().replace(/\s+/g, ' ');
        if (trimmed !== original) {
          count++;
          if (preview) {
            range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
          } else {
            data[r][c] = trimmed;
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

/**
 * Fixes capitalization in scope. Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function fixCapitalization_(preview) {
  const settings = loadSettings();
  const style = settings.capStyle || 'title';
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === 'string' && data[r][c].trim()) {
        const original = data[r][c];
        let fixed;

        switch (style) {
          case 'upper':
            fixed = original.toUpperCase();
            break;
          case 'lower':
            fixed = original.toLowerCase();
            break;
          case 'sentence':
            fixed = toSentenceCase_(original);
            break;
          case 'title':
          default:
            fixed = toTitleCase_(original);
            break;
        }

        if (fixed !== original) {
          count++;
          if (preview) {
            range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
          } else {
            data[r][c] = fixed;
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

/**
 * Standardizes dates in scope to a consistent format. Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function standardizeDates_(preview) {
  const settings = loadSettings();
  const targetFormat = settings.dateFormat || 'MM/DD/YYYY';
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = data[r][c];
      const parsed = parseDate_(val);

      if (parsed) {
        const formatted = formatDateAs_(parsed, targetFormat);
        const currentStr = String(val);

        if (formatted !== currentStr) {
          count++;
          if (preview) {
            range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
          } else {
            data[r][c] = formatted;
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

/**
 * Removes empty rows and columns. Returns { rows, cols } removed.
 * @param {boolean} preview — if true, highlights empties instead of removing
 */
function removeEmpties_(preview) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  let rowCount = 0;
  let colCount = 0;

  // Find empty rows (skip header row 0)
  const emptyRows = [];
  for (let r = data.length - 1; r >= 1; r--) {
    const isEmpty = data[r].every(cell => cell === '' || cell === null || cell === undefined);
    if (isEmpty) {
      emptyRows.push(r);
      rowCount++;
    }
  }

  // Find empty columns
  const emptyCols = [];
  if (data.length > 0) {
    for (let c = data[0].length - 1; c >= 0; c--) {
      const isEmpty = data.every(row => row[c] === '' || row[c] === null || row[c] === undefined);
      if (isEmpty) {
        emptyCols.push(c);
        colCount++;
      }
    }
  }

  if (preview) {
    for (const r of emptyRows) {
      sheet.getRange(r + 1, 1, 1, data[0].length).setBackground('#FFF9C4');
    }
    for (const c of emptyCols) {
      sheet.getRange(1, c + 1, data.length, 1).setBackground('#FFF9C4');
    }
  } else {
    // Delete rows from bottom to top
    for (const r of emptyRows) {
      sheet.deleteRow(r + 1);
    }
    // Delete columns from right to left
    for (const c of emptyCols) {
      sheet.deleteColumn(c + 1);
    }
  }

  return { rows: rowCount, cols: colCount };
}

/**
 * Standardizes phone numbers. Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function standardizePhones_(preview) {
  const settings = loadSettings();
  const format = settings.phoneFormat || '(XXX) XXX-XXXX';
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = String(data[r][c]);
      const digits = val.replace(/\D/g, '');

      // Must look like a phone number (7, 10, or 11 digits)
      if (digits.length >= 7 && digits.length <= 11 && /\d/.test(val)) {
        // Use the last 10 digits (strip country code if 11)
        let d = digits;
        if (d.length === 11 && d.startsWith('1')) {
          d = d.substring(1);
        }
        if (d.length === 10) {
          let formatted;
          switch (format) {
            case 'XXX-XXX-XXXX':
              formatted = d.substring(0, 3) + '-' + d.substring(3, 6) + '-' + d.substring(6);
              break;
            case 'XXX.XXX.XXXX':
              formatted = d.substring(0, 3) + '.' + d.substring(3, 6) + '.' + d.substring(6);
              break;
            case '+1 (XXX) XXX-XXXX':
              formatted = '+1 (' + d.substring(0, 3) + ') ' + d.substring(3, 6) + '-' + d.substring(6);
              break;
            case '(XXX) XXX-XXXX':
            default:
              formatted = '(' + d.substring(0, 3) + ') ' + d.substring(3, 6) + '-' + d.substring(6);
              break;
          }

          if (formatted !== val) {
            count++;
            if (preview) {
              range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
            } else {
              data[r][c] = formatted;
            }
          }
        } else if (d.length === 7) {
          const formatted = d.substring(0, 3) + '-' + d.substring(3);
          if (formatted !== val) {
            count++;
            if (preview) {
              range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
            } else {
              data[r][c] = formatted;
            }
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

/**
 * Fixes email formatting (lowercase + trim). Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function fixEmails_(preview) {
  const settings = loadSettings();
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === 'string') {
        const original = data[r][c];
        const trimmed = original.trim();

        if (emailRegex.test(trimmed)) {
          const fixed = trimmed.toLowerCase();
          if (fixed !== original) {
            count++;
            if (preview) {
              range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
            } else {
              data[r][c] = fixed;
            }
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

/**
 * Removes special characters from cells. Returns count of cells modified.
 * @param {boolean} preview — if true, highlights cells instead of modifying
 */
function removeSpecialChars_(preview) {
  const settings = loadSettings();
  const keepChars = settings.specialCharsKeep || '.,!?@#$%&()-';
  const range = getScopeRange_(settings);
  const data = range.getValues();
  let count = 0;

  // Build regex: keep alphanumerics, spaces, and the allowed special chars
  const escaped = keepChars.replace(/[-[\]{}()*+?.,\\^$|#]/g, '\\$&');
  const pattern = new RegExp('[^a-zA-Z0-9\\s' + escaped + ']', 'g');

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === 'string' && data[r][c].trim()) {
        const original = data[r][c];
        const cleaned = original.replace(pattern, '');
        if (cleaned !== original) {
          count++;
          if (preview) {
            range.getCell(r + 1, c + 1).setBackground('#FFF9C4');
          } else {
            data[r][c] = cleaned;
          }
        }
      }
    }
  }

  if (!preview && count > 0) {
    range.setValues(data);
  }

  return count;
}

// ═══════════════════════════════════════════
// TEXT HELPERS
// ═══════════════════════════════════════════

function toTitleCase_(str) {
  const lowers = ['a','an','the','and','but','or','for','nor','on','at','to','by','in','of','up','as','is','it'];
  return str.replace(/\w\S*/g, function (word, index) {
    if (index > 0 && lowers.includes(word.toLowerCase())) {
      return word.toLowerCase();
    }
    return word.charAt(0).toUpperCase() + word.substring(1).toLowerCase();
  });
}

function toSentenceCase_(str) {
  return str.charAt(0).toUpperCase() + str.substring(1).toLowerCase();
}

// ═══════════════════════════════════════════
// DATE HELPERS
// ═══════════════════════════════════════════

/**
 * Attempts to parse a value as a date. Returns a Date object or null.
 */
function parseDate_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val;
  }

  if (typeof val !== 'string') return null;

  const s = val.trim();
  if (!s) return null;

  // Common date patterns
  const patterns = [
    // MM/DD/YYYY or MM-DD-YYYY
    { regex: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/, parse: function(m) { return new Date(+m[3], +m[1]-1, +m[2]); } },
    // YYYY-MM-DD
    { regex: /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/, parse: function(m) { return new Date(+m[1], +m[2]-1, +m[3]); } },
    // DD/MM/YYYY (try if month > 12 not possible — handled by checking)
    // Month DD, YYYY
    { regex: /^([A-Za-z]+)\s+(\d{1,2}),?\s*(\d{4})$/, parse: function(m) { var mo = monthIndex_(m[1]); return mo >= 0 ? new Date(+m[3], mo, +m[2]) : null; } },
    // DD Month YYYY
    { regex: /^(\d{1,2})\s+([A-Za-z]+),?\s*(\d{4})$/, parse: function(m) { var mo = monthIndex_(m[2]); return mo >= 0 ? new Date(+m[3], mo, +m[1]) : null; } },
    // M/D/YY
    { regex: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$/, parse: function(m) { var yr = +m[3] > 50 ? 1900 + +m[3] : 2000 + +m[3]; return new Date(yr, +m[1]-1, +m[2]); } },
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = s.match(patterns[i].regex);
    if (match) {
      var d = patterns[i].parse(match);
      if (d && !isNaN(d.getTime())) return d;
    }
  }

  return null;
}

function monthIndex_(name) {
  var months = {
    jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,
    may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,
    oct:9,october:9,nov:10,november:10,dec:11,december:11
  };
  return months[name.toLowerCase()] !== undefined ? months[name.toLowerCase()] : -1;
}

/**
 * Formats a Date object to the target format string.
 */
function formatDateAs_(date, format) {
  var mm = String(date.getMonth() + 1).padStart(2, '0');
  var dd = String(date.getDate()).padStart(2, '0');
  var yyyy = String(date.getFullYear());

  switch (format) {
    case 'YYYY-MM-DD':
      return yyyy + '-' + mm + '-' + dd;
    case 'DD/MM/YYYY':
      return dd + '/' + mm + '/' + yyyy;
    case 'Month DD, YYYY':
      var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      return months[date.getMonth()] + ' ' + date.getDate() + ', ' + yyyy;
    case 'MM/DD/YYYY':
    default:
      return mm + '/' + dd + '/' + yyyy;
  }
}

// ═══════════════════════════════════════════
// MENU ENTRY POINTS
// ═══════════════════════════════════════════

function menuRemoveDuplicates() {
  createBackup_();
  var count = removeDuplicates_(false);
  showResultsDialog_({ duplicates: count });
}

function menuTrimWhitespace() {
  createBackup_();
  var count = trimWhitespace_(false);
  showResultsDialog_({ trimmed: count });
}

function menuFixCapitalization() {
  createBackup_();
  var count = fixCapitalization_(false);
  showResultsDialog_({ capitalization: count });
}

function menuStandardizeDates() {
  createBackup_();
  var count = standardizeDates_(false);
  showResultsDialog_({ dates: count });
}

function menuRemoveEmpties() {
  createBackup_();
  var result = removeEmpties_(false);
  showResultsDialog_({ emptyRows: result.rows, emptyCols: result.cols });
}

function menuStandardizePhones() {
  createBackup_();
  var count = standardizePhones_(false);
  showResultsDialog_({ phones: count });
}

function menuFixEmails() {
  createBackup_();
  var count = fixEmails_(false);
  showResultsDialog_({ emails: count });
}

function menuRemoveSpecialChars() {
  createBackup_();
  var count = removeSpecialChars_(false);
  showResultsDialog_({ specialChars: count });
}

// ═══════════════════════════════════════════
// RUN ALL & PREVIEW
// ═══════════════════════════════════════════

/**
 * Runs all cleanup operations in sequence.
 */
function runAllCleanups() {
  createBackup_();
  var results = runAllOperations_(false);
  showResultsDialog_(results);
}

/**
 * Preview mode — highlights cells that would change in yellow.
 */
function previewChanges() {
  // Clear any previous preview highlights
  clearPreviewHighlights_();
  var results = runAllOperations_(true);
  showPreviewDialog_(results);
}

/**
 * Runs all operations. Called by both run and preview.
 * @param {boolean} preview
 */
function runAllOperations_(preview) {
  var dups = removeDuplicates_(preview);
  var trimmed = trimWhitespace_(preview);
  var caps = fixCapitalization_(preview);
  var dates = standardizeDates_(preview);
  var empties = removeEmpties_(preview);
  var phones = standardizePhones_(preview);
  var emails = fixEmails_(preview);
  var specials = removeSpecialChars_(preview);

  return {
    duplicates: dups,
    trimmed: trimmed,
    capitalization: caps,
    dates: dates,
    emptyRows: empties.rows,
    emptyCols: empties.cols,
    phones: phones,
    emails: emails,
    specialChars: specials,
  };
}

/**
 * Runs all cleanups from the sidebar with specific options checked.
 */
function runFromSidebar(options) {
  createBackup_();
  var results = {};

  if (options.duplicates) results.duplicates = removeDuplicates_(false);
  if (options.whitespace) results.trimmed = trimWhitespace_(false);
  if (options.capitalization) results.capitalization = fixCapitalization_(false);
  if (options.dates) results.dates = standardizeDates_(false);
  if (options.empties) {
    var e = removeEmpties_(false);
    results.emptyRows = e.rows;
    results.emptyCols = e.cols;
  }
  if (options.phones) results.phones = standardizePhones_(false);
  if (options.emails) results.emails = fixEmails_(false);
  if (options.specialChars) results.specialChars = removeSpecialChars_(false);

  logCleanupRun_(results, 'Run');
  return results;
}

/**
 * Runs preview from the sidebar with specific options checked.
 */
function previewFromSidebar(options) {
  clearPreviewHighlights_();
  var results = {};

  if (options.duplicates) results.duplicates = removeDuplicates_(true);
  if (options.whitespace) results.trimmed = trimWhitespace_(true);
  if (options.capitalization) results.capitalization = fixCapitalization_(true);
  if (options.dates) results.dates = standardizeDates_(true);
  if (options.empties) {
    var e = removeEmpties_(true);
    results.emptyRows = e.rows;
    results.emptyCols = e.cols;
  }
  if (options.phones) results.phones = standardizePhones_(true);
  if (options.emails) results.emails = fixEmails_(true);
  if (options.specialChars) results.specialChars = removeSpecialChars_(true);

  return results;
}

/**
 * Clears yellow preview highlights from the active sheet.
 */
function clearPreviewHighlights_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var backgrounds = range.getBackgrounds();
  var changed = false;

  for (var r = 0; r < backgrounds.length; r++) {
    for (var c = 0; c < backgrounds[r].length; c++) {
      if (backgrounds[r][c] === '#fff9c4' || backgrounds[r][c] === '#FFF9C4') {
        backgrounds[r][c] = null; // Reset to default
        changed = true;
      }
    }
  }

  if (changed) {
    range.setBackgrounds(backgrounds);
  }
}

// ═══════════════════════════════════════════
// RESULTS DIALOGS
// ═══════════════════════════════════════════

/**
 * Shows a styled results dialog after cleanup.
 */
function showResultsDialog_(results) {
  var lines = [];

  if (results.duplicates !== undefined)    lines.push(statusLine_(results.duplicates, 'duplicate rows removed'));
  if (results.trimmed !== undefined)       lines.push(statusLine_(results.trimmed, 'cells trimmed'));
  if (results.capitalization !== undefined) lines.push(statusLine_(results.capitalization, 'cells re-capitalized'));
  if (results.dates !== undefined)         lines.push(statusLine_(results.dates, 'dates standardized'));
  if (results.emptyRows !== undefined)     lines.push(statusLine_(results.emptyRows, 'empty rows removed'));
  if (results.emptyCols !== undefined)     lines.push(statusLine_(results.emptyCols, 'empty columns removed'));
  if (results.phones !== undefined)        lines.push(statusLine_(results.phones, 'phone numbers formatted'));
  if (results.emails !== undefined)        lines.push(statusLine_(results.emails, 'emails fixed'));
  if (results.specialChars !== undefined)  lines.push(statusLine_(results.specialChars, 'cells cleaned of special chars'));

  var total = Object.values(results).reduce(function(sum, v) { return sum + (v || 0); }, 0);

  var html = '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px;">' +
    '<div style="text-align: center; margin-bottom: 16px;">' +
    '<div style="font-size: 28px;">🕷</div>' +
    '<h2 style="margin: 4px 0; font-size: 16px;">Cleanup Complete</h2>' +
    '<p style="color: #666; font-size: 12px;">A backup sheet was created before changes.</p>' +
    '</div>' +
    '<div style="background: #f5f5f5; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;">' +
    lines.join('') +
    '</div>' +
    '<div style="text-align: center; padding: 10px; border-radius: 8px; ' +
    (total > 0 ? 'background: #E8F5E9; color: #2E7D32;' : 'background: #f5f5f5; color: #666;') + '">' +
    '<strong>' + (total > 0 ? total + ' total changes made' : 'No changes needed — data is clean!') + '</strong>' +
    '</div></div>';

  logCleanupRun_(results, 'Run');

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(360).setHeight(380),
    'Cleanup Results'
  );
}

/**
 * Shows a preview results dialog.
 */
function showPreviewDialog_(results) {
  var total = Object.values(results).reduce(function(sum, v) { return sum + (v || 0); }, 0);

  var html = '<div style="font-family: \'Segoe UI\', system-ui, sans-serif; padding: 20px; text-align: center;">' +
    '<div style="font-size: 28px;">🧪</div>' +
    '<h2 style="margin: 4px 0 12px; font-size: 16px;">Preview Complete</h2>' +
    '<p style="font-size: 13px; color: #333; line-height: 1.6;">' +
    (total > 0
      ? '<strong>' + total + ' cells</strong> would be changed.<br>They are highlighted in <span style="background:#FFF9C4;padding:2px 6px;border-radius:3px;">yellow</span> on your sheet.'
      : 'No changes needed — your data looks clean!') +
    '</p>' +
    '<p style="font-size: 12px; color: #999; margin-top: 12px;">No data was modified. Run a cleanup to apply changes.</p>' +
    '</div>';

  logCleanupRun_(results, 'Preview');

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(320).setHeight(220),
    'Preview Results'
  );
}

function statusLine_(count, label) {
  var color = count > 0 ? '#2E7D32' : '#999';
  var icon = count > 0 ? '✓' : '—';
  return '<div style="padding: 4px 0; font-size: 13px; color: ' + color + ';">' +
    '<span style="margin-right: 6px;">' + icon + '</span>' +
    '<strong>' + count + '</strong> ' + label + '</div>';
}

// ═══════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════

/**
 * Gets or creates the Cleanup Dashboard sheet with branded stats bar.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);

  // Row 1: stat values (gold, large)
  sheet.getRange(1, 1, 1, 5)
    .setValues([['—', '—', '—', '—', '—']])
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  // Row 2: stat labels
  sheet.getRange(2, 1, 1, 5)
    .setValues([['TOTAL RUNS', 'CELLS FIXED', 'DUPES REMOVED', 'PREVIEWS', 'LAST RUN']])
    .setBackground('#1A1A1A')
    .setFontColor('#888888')
    .setFontFamily('Roboto Mono')
    .setFontSize(8)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top');
  sheet.setRowHeight(2, 20);

  // Row 3: column headers
  const headers = ['Timestamp', 'Sheet', 'Operation', 'Changes Made', 'Mode'];
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#1A1A1A')
    .setFontColor('#C9A84C')
    .setFontFamily('Roboto Mono')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);

  // Column widths
  sheet.setColumnWidth(1, 160);  // Timestamp
  sheet.setColumnWidth(2, 180);  // Sheet
  sheet.setColumnWidth(3, 200);  // Operation
  sheet.setColumnWidth(4, 130);  // Changes Made
  sheet.setColumnWidth(5, 100);  // Mode

  return sheet;
}

/**
 * Reads log data and updates the stats bar (row 1).
 */
function refreshDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  const dataStartRow = DASHBOARD_HEADER_ROW + 1;
  const lastRow = sheet.getLastRow();

  let totalRuns = 0;
  let cellsFixed = 0;
  let dupesRemoved = 0;
  let previews = 0;
  let lastRunTs = null;

  if (lastRow >= dataStartRow) {
    const data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 5).getValues();

    for (const row of data) {
      const ts = row[0];
      const operation = String(row[2]);
      const changes = Number(row[3]) || 0;
      const mode = String(row[4]);

      if (mode === 'Run') {
        totalRuns++;
        cellsFixed += changes;
        if (operation === 'Remove Duplicates' || operation === 'All Cleanups') {
          // can't distinguish dupes from total in All Cleanups — only count dedicated runs
          if (operation === 'Remove Duplicates') dupesRemoved += changes;
        }
      } else if (mode === 'Preview') {
        previews++;
      }

      if (ts instanceof Date && (!lastRunTs || ts > lastRunTs)) lastRunTs = ts;
    }
  }

  const lastRunDisplay = lastRunTs
    ? Utilities.formatDate(lastRunTs, Session.getScriptTimeZone(), 'MMM d')
    : '—';

  sheet.getRange(1, 1, 1, 5).setValues([[
    totalRuns || '—',
    cellsFixed || '—',
    dupesRemoved || '—',
    previews || '—',
    lastRunDisplay,
  ]]);
}

/**
 * Derives a human-readable operation name from a results object.
 * @param {Object} results
 * @return {string}
 */
function getOperationName_(results) {
  const definedKeys = Object.keys(results).filter(k => results[k] !== undefined);
  if (definedKeys.length >= 4) return 'All Cleanups';
  const nameMap = {
    duplicates: 'Remove Duplicates',
    trimmed: 'Trim Whitespace',
    capitalization: 'Fix Capitalization',
    dates: 'Standardize Dates',
    emptyRows: 'Remove Empties',
    emptyCols: 'Remove Empties',
    phones: 'Standardize Phones',
    emails: 'Fix Emails',
    specialChars: 'Remove Special Chars',
  };
  const first = definedKeys.find(k => nameMap[k]);
  return first ? nameMap[first] : 'Cleanup';
}

/**
 * Logs a cleanup run to the Cleanup Dashboard sheet.
 * @param {Object} results - The results object from the cleanup.
 * @param {string} mode - 'Run' or 'Preview'.
 */
function logCleanupRun_(results, mode) {
  try {
    const sheet = getOrCreateDashboard_();
    const dataStartRow = DASHBOARD_HEADER_ROW + 1;
    const activeSheetName = SpreadsheetApp.getActiveSheet().getName();
    const operation = getOperationName_(results);
    const total = Object.values(results).reduce((sum, v) => sum + (Number(v) || 0), 0);

    const targetRow = Math.max(sheet.getLastRow() + 1, dataStartRow);
    sheet.getRange(targetRow, 1, 1, 5).setValues([[
      new Date(), activeSheetName, operation, total, mode,
    ]]);

    // Style the mode cell
    const modeCell = sheet.getRange(targetRow, 5);
    if (mode === 'Run') {
      modeCell.setBackground('#E8F5E9').setFontColor('#2E7D32').setFontWeight('bold');
    } else {
      modeCell.setBackground('#FFF8E1').setFontColor('#F57F17').setFontWeight('bold');
    }

    // Style changes cell — gold if > 0
    const changesCell = sheet.getRange(targetRow, 4);
    if (total > 0) {
      changesCell.setFontColor('#C9A84C').setFontWeight('bold');
    }

    // Alternate row background (preserve status cell colors)
    const rowRange = sheet.getRange(targetRow, 1, 1, 3);
    rowRange.setBackground(targetRow % 2 === 0 ? '#F9F9F9' : '#FFFFFF')
      .setFontFamily('Roboto').setFontSize(10);

    refreshDashboardStats();
  } catch (e) {
    Logger.log('Dashboard log error: ' + e.message);
  }
}

// ═══════════════════════════════════════════
// SIDEBAR HTML
// ═══════════════════════════════════════════

function getSidebarHtml() {
  return '<!DOCTYPE html>\
<html>\
<head>\
  <style>\
    * { box-sizing: border-box; margin: 0; padding: 0; }\
    body {\
      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;\
      background: #fafafa;\
      color: #1a1a1a;\
      font-size: 13px;\
    }\
    .header {\
      background: #1a1a1a;\
      color: white;\
      padding: 20px 16px;\
      text-align: center;\
    }\
    .header .logo { font-size: 24px; margin-bottom: 4px; }\
    .header h1 { font-size: 15px; font-weight: 600; letter-spacing: 0.5px; }\
    .header .brand { color: #C9A84C; }\
    .header .sub { font-size: 11px; color: #888; margin-top: 4px; }\
    .section {\
      padding: 16px;\
      border-bottom: 1px solid #eee;\
    }\
    .section-title {\
      font-size: 11px;\
      font-weight: 600;\
      text-transform: uppercase;\
      letter-spacing: 1px;\
      color: #666;\
      margin-bottom: 10px;\
    }\
    .checkbox-row {\
      display: flex;\
      align-items: center;\
      padding: 7px 0;\
    }\
    .checkbox-row input[type="checkbox"] {\
      margin-right: 10px;\
      width: 16px;\
      height: 16px;\
      accent-color: #C9A84C;\
      cursor: pointer;\
    }\
    .checkbox-row label {\
      font-size: 13px;\
      cursor: pointer;\
      flex: 1;\
    }\
    .checkbox-row .icon {\
      margin-right: 6px;\
    }\
    .field { margin-bottom: 14px; }\
    .field label {\
      display: block;\
      font-size: 11px;\
      font-weight: 600;\
      text-transform: uppercase;\
      letter-spacing: 1px;\
      color: #666;\
      margin-bottom: 6px;\
    }\
    .field select, .field input {\
      width: 100%;\
      padding: 9px 12px;\
      border: 1px solid #ddd;\
      border-radius: 6px;\
      font-size: 13px;\
      font-family: inherit;\
      background: white;\
      transition: border-color 0.2s;\
    }\
    .field select:focus, .field input:focus {\
      outline: none;\
      border-color: #C9A84C;\
      box-shadow: 0 0 0 3px rgba(201,168,76,0.1);\
    }\
    .field .help {\
      font-size: 11px;\
      color: #999;\
      margin-top: 4px;\
      line-height: 1.4;\
    }\
    .btn-group { padding: 16px; }\
    .btn {\
      width: 100%;\
      padding: 12px;\
      border: none;\
      border-radius: 8px;\
      font-size: 13px;\
      font-weight: 600;\
      cursor: pointer;\
      transition: all 0.2s;\
      letter-spacing: 0.5px;\
    }\
    .btn-primary {\
      background: #1a1a1a;\
      color: #C9A84C;\
      border: 1px solid #C9A84C;\
    }\
    .btn-primary:hover { background: #C9A84C; color: #1a1a1a; }\
    .btn-secondary {\
      background: white;\
      color: #666;\
      border: 1px solid #ddd;\
      margin-top: 8px;\
    }\
    .btn-secondary:hover { border-color: #999; color: #333; }\
    .btn-text {\
      background: none;\
      color: #999;\
      border: none;\
      margin-top: 8px;\
      font-size: 12px;\
    }\
    .btn-text:hover { color: #666; }\
    .status {\
      text-align: center;\
      padding: 8px;\
      font-size: 12px;\
      margin-top: 10px;\
      border-radius: 6px;\
      display: none;\
    }\
    .status.success { display: block; background: #E8F5E9; color: #2E7D32; }\
    .status.error { display: block; background: #FFEBEE; color: #C62828; }\
    .status.warning { display: block; background: #FFF8E1; color: #F57F17; }\
    .select-all-row {\
      display: flex;\
      justify-content: space-between;\
      align-items: center;\
      margin-bottom: 6px;\
    }\
    .select-all-row a {\
      font-size: 11px;\
      color: #C9A84C;\
      cursor: pointer;\
      text-decoration: none;\
    }\
    .select-all-row a:hover { text-decoration: underline; }\
  </style>\
</head>\
<body>\
  <div class="header">\
    <div class="logo">🕷</div>\
    <h1><span class="brand">TAK</span>Scripts</h1>\
    <div class="sub">Data Cleanup Wizard</div>\
  </div>\
\
  <div class="section">\
    <div class="select-all-row">\
      <div class="section-title">Cleanup Operations</div>\
      <div><a onclick="selectAll()">Select All</a> · <a onclick="selectNone()">None</a></div>\
    </div>\
    <div class="checkbox-row"><input type="checkbox" id="chkDuplicates" checked /><label for="chkDuplicates"><span class="icon">🔁</span> Remove Duplicates</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkWhitespace" checked /><label for="chkWhitespace"><span class="icon">✂️</span> Trim Whitespace</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkCapitalization" /><label for="chkCapitalization"><span class="icon">🔤</span> Fix Capitalization</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkDates" /><label for="chkDates"><span class="icon">📅</span> Standardize Dates</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkEmpties" checked /><label for="chkEmpties"><span class="icon">🗑</span> Remove Empty Rows/Cols</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkPhones" /><label for="chkPhones"><span class="icon">📞</span> Standardize Phones</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkEmails" checked /><label for="chkEmails"><span class="icon">📧</span> Fix Email Formatting</label></div>\
    <div class="checkbox-row"><input type="checkbox" id="chkSpecialChars" /><label for="chkSpecialChars"><span class="icon">🚫</span> Remove Special Chars</label></div>\
  </div>\
\
  <div class="section">\
    <div class="section-title">Scope</div>\
    <div class="field">\
      <select id="scope">\
        <option value="sheet">Entire Sheet</option>\
        <option value="selection">Selected Range</option>\
        <option value="columns">Specific Columns</option>\
      </select>\
    </div>\
    <div class="field" id="colField" style="display:none;">\
      <label>Columns</label>\
      <input type="text" id="scopeColumns" placeholder="A, B, D" />\
      <div class="help">Comma-separated column letters.</div>\
    </div>\
  </div>\
\
  <div class="section">\
    <div class="section-title">Options</div>\
    <div class="field">\
      <label>Date Format</label>\
      <select id="dateFormat">\
        <option value="MM/DD/YYYY">MM/DD/YYYY</option>\
        <option value="YYYY-MM-DD">YYYY-MM-DD</option>\
        <option value="DD/MM/YYYY">DD/MM/YYYY</option>\
        <option value="Month DD, YYYY">Month DD, YYYY</option>\
      </select>\
    </div>\
    <div class="field">\
      <label>Phone Format</label>\
      <select id="phoneFormat">\
        <option value="(XXX) XXX-XXXX">(XXX) XXX-XXXX</option>\
        <option value="XXX-XXX-XXXX">XXX-XXX-XXXX</option>\
        <option value="XXX.XXX.XXXX">XXX.XXX.XXXX</option>\
        <option value="+1 (XXX) XXX-XXXX">+1 (XXX) XXX-XXXX</option>\
      </select>\
    </div>\
    <div class="field">\
      <label>Capitalization Style</label>\
      <select id="capStyle">\
        <option value="title">Title Case</option>\
        <option value="upper">UPPERCASE</option>\
        <option value="lower">lowercase</option>\
        <option value="sentence">Sentence case</option>\
      </select>\
    </div>\
    <div class="field">\
      <label>Duplicate Check Columns</label>\
      <input type="text" id="dupColumns" placeholder="Leave blank for all columns" />\
      <div class="help">Comma-separated column letters (e.g. A, B). Blank = compare all.</div>\
    </div>\
    <div class="field">\
      <label>Keep Special Characters</label>\
      <input type="text" id="specialCharsKeep" placeholder=".,!?@#$%&()-" />\
      <div class="help">Characters to preserve when removing specials.</div>\
    </div>\
  </div>\
\
  <div class="btn-group">\
    <button class="btn btn-primary" onclick="runCleanup()">🧹 Run Cleanup</button>\
    <button class="btn btn-secondary" onclick="runPreview()">🧪 Preview Changes</button>\
    <button class="btn btn-text" onclick="saveSettingsOnly()">💾 Save Settings</button>\
    <div id="status" class="status"></div>\
  </div>\
\
  <script>\
    document.getElementById("scope").addEventListener("change", function() {\
      document.getElementById("colField").style.display = this.value === "columns" ? "block" : "none";\
    });\
\
    google.script.run.withSuccessHandler(function(s) {\
      if (s.scope) document.getElementById("scope").value = s.scope;\
      if (s.dateFormat) document.getElementById("dateFormat").value = s.dateFormat;\
      if (s.phoneFormat) document.getElementById("phoneFormat").value = s.phoneFormat;\
      if (s.capStyle) document.getElementById("capStyle").value = s.capStyle;\
      if (s.dupColumns) document.getElementById("dupColumns").value = s.dupColumns;\
      if (s.specialCharsKeep) document.getElementById("specialCharsKeep").value = s.specialCharsKeep;\
      if (s.scopeColumns) document.getElementById("scopeColumns").value = s.scopeColumns;\
      if (s.scope === "columns") document.getElementById("colField").style.display = "block";\
    }).loadSettings();\
\
    function selectAll() {\
      document.querySelectorAll("input[type=checkbox]").forEach(function(cb) { cb.checked = true; });\
    }\
    function selectNone() {\
      document.querySelectorAll("input[type=checkbox]").forEach(function(cb) { cb.checked = false; });\
    }\
\
    function getSettings() {\
      return {\
        scope: document.getElementById("scope").value,\
        scopeColumns: document.getElementById("scopeColumns").value,\
        dateFormat: document.getElementById("dateFormat").value,\
        phoneFormat: document.getElementById("phoneFormat").value,\
        capStyle: document.getElementById("capStyle").value,\
        dupColumns: document.getElementById("dupColumns").value,\
        specialCharsKeep: document.getElementById("specialCharsKeep").value,\
      };\
    }\
\
    function getOptions() {\
      return {\
        duplicates: document.getElementById("chkDuplicates").checked,\
        whitespace: document.getElementById("chkWhitespace").checked,\
        capitalization: document.getElementById("chkCapitalization").checked,\
        dates: document.getElementById("chkDates").checked,\
        empties: document.getElementById("chkEmpties").checked,\
        phones: document.getElementById("chkPhones").checked,\
        emails: document.getElementById("chkEmails").checked,\
        specialChars: document.getElementById("chkSpecialChars").checked,\
      };\
    }\
\
    function showStatus(msg, type) {\
      var el = document.getElementById("status");\
      el.textContent = msg;\
      el.className = "status " + type;\
    }\
\
    function runCleanup() {\
      showStatus("Running cleanup...", "warning");\
      google.script.run\
        .withSuccessHandler(function() { showStatus("Saving settings...", "warning"); })\
        .saveSettings(getSettings());\
      setTimeout(function() {\
        google.script.run\
          .withSuccessHandler(function(results) {\
            var total = 0;\
            for (var k in results) total += results[k] || 0;\
            showStatus(total > 0 ? "Done! " + total + " changes made." : "No changes needed — data is clean!", "success");\
          })\
          .withFailureHandler(function(err) {\
            showStatus("Error: " + err.message, "error");\
          })\
          .runFromSidebar(getOptions());\
      }, 500);\
    }\
\
    function runPreview() {\
      showStatus("Generating preview...", "warning");\
      google.script.run\
        .withSuccessHandler(function() { })\
        .saveSettings(getSettings());\
      setTimeout(function() {\
        google.script.run\
          .withSuccessHandler(function(results) {\
            var total = 0;\
            for (var k in results) total += results[k] || 0;\
            showStatus(total > 0 ? total + " cells highlighted in yellow." : "No changes needed!", "success");\
          })\
          .withFailureHandler(function(err) {\
            showStatus("Error: " + err.message, "error");\
          })\
          .previewFromSidebar(getOptions());\
      }, 500);\
    }\
\
    function saveSettingsOnly() {\
      google.script.run\
        .withSuccessHandler(function() {\
          showStatus("Settings saved.", "success");\
        })\
        .withFailureHandler(function(err) {\
          showStatus("Error: " + err.message, "error");\
        })\
        .saveSettings(getSettings());\
    }\
  </script>\
</body>\
</html>';
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
  '<h2>Data Cleanup Wizard</h2><p>Quick Reference Guide</p></div>' +
  '<div class="content">' +
  '<div class="section"><h3>Quick Start</h3><ol>' +
  '<li>Click on the sheet tab you want to clean, or select a specific range</li>' +
  '<li>Click <strong>▶️ Preview Cleanup</strong> to see exactly what would change — no data is touched yet</li>' +
  '<li>Review the preview, then choose your cleanup operations</li>' +
  '<li>Click <strong>Run Cleanup</strong> to apply the changes</li>' +
  '<li>The dashboard logs every run with a count of changes made</li>' +
  '</ol></div>' +
  '<div class="section"><h3>Cleanup Operations</h3>' +
  '<div class="setting"><strong>Trim Whitespace</strong><span>Removes leading and trailing spaces from every cell. Fixes copy-paste issues.</span></div>' +
  '<div class="setting"><strong>Fix Text Case</strong><span>Converts text to Title Case, UPPER, lower, or Sentence case consistently</span></div>' +
  '<div class="setting"><strong>Remove Duplicates</strong><span>Finds and removes duplicate rows based on a column you choose</span></div>' +
  '<div class="setting"><strong>Remove Blank Rows</strong><span>Deletes completely empty rows from your data range</span></div>' +
  '</div>' +
  '<div class="section"><h3>Tips</h3>' +
  '<div class="tip">Always <strong>Preview first</strong> — see the exact changes before committing anything</div>' +
  '<div class="tip">Enable <strong>Backup Before Cleaning</strong> in settings for a safety net on important data</div>' +
  '<div class="tip">Use <strong>Undo</strong> from the menu to restore from the last backup if something looks wrong</div>' +
  '</div></div>' +
  '<div class="footer">TAKScripts · takscripts.store</div>' +
  '</body></html>';
}
