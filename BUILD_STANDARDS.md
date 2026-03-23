# TAKScripts Build Standards

Rules and patterns established across the first 11 scripts.
Apply these to every new script from #12 onwards.

---

## 1. Settings Save Pattern

**Rule:** Always save to `PropertiesService` first. Wrap any `ScriptApp` trigger calls in a try/catch after the save.

**Why:** `ScriptApp` requires a scope that may not be authorized on first use. If it throws before `setProperty()` is called, settings are lost and the sidebar shows a false error.

```javascript
function saveSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(SETTINGS_KEY, JSON.stringify(settings)); // ← save FIRST

  try {
    updateScheduleTrigger_(settings.autoSchedule); // ← trigger AFTER, in try/catch
  } catch (e) {
    Logger.log('Trigger update skipped: ' + e.message);
  }

  return { success: true };
}
```

**Applies to:** Any script with auto-schedule functionality (inventory, auto-file-organizer, attachment-auto-saver, follow-up-nudger, etc.)

---

## 2. Dashboard Layout Pattern

**Rule:** Every script gets a 3-row stats bar at the top of its primary sheet. No exceptions.

```
Row 1 — Stat values    (gold #C9A84C, 20pt Roboto Mono, dark bg #1A1A1A, 48px tall)
Row 2 — Stat labels    (8pt, #888888, dark bg, 20px tall)
Row 3 — Column headers (gold on dark, 9pt Roboto Mono, bold)
Row 4+ — Data
```

**Key constants (copy into every script):**
```javascript
var DASHBOARD_SHEET_NAME = '📊 [Name] Dashboard';
var DASHBOARD_HEADER_ROW = 3;
```

**Key functions every script must have:**
- `getOrCreateDashboard_()` — idempotent sheet creation, handles migration from old sheet names
- `refreshDashboardStats()` — reads live data, writes values to row 1

**Rules:**
- `setFrozenRows(DASHBOARD_HEADER_ROW)` — always freeze stats + headers
- `Math.max(sheet.getLastRow() + 1, dataStartRow)` — always use safe row insertion
- **No footers** — never call `addSheetFooter_()` or append footer rows
- For live-view sheets (cleared and rewritten each run): clear rows `>= DASHBOARD_HEADER_ROW + 1` only, never `sheet.clear()`

---

## 3. Multiple Input Fields (Comma-Separated → OR Query)

**Rule:** Whenever a script filters by sender, label, keyword, or any other list — accept comma-separated input and build the query with OR logic. Never limit to a single value.

**Why:** Customers always need to filter by multiple senders/labels in real use.

**Pattern (Gmail query):**
```javascript
// In getDefaultSettings_():
senders: '',  // comma-separated email addresses

// In buildSearchQuery_():
if (settings.senders && settings.senders.trim()) {
  var senderList = settings.senders.split(',')
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s.length > 0; });
  if (senderList.length === 1) {
    parts.push('from:' + senderList[0]);
  } else if (senderList.length > 1) {
    parts.push('(' + senderList.map(function(s) { return 'from:' + s; }).join(' OR ') + ')');
  }
}
```

**Sidebar HTML field:**
```html
<textarea id="senders" rows="3"
  placeholder="accounting@acme.com, invoices@supplier.com">
</textarea>
<div class="help">Comma-separated. Leave blank for all senders.</div>
```

**Applies to:** Any script that scans Gmail by sender, label, keyword, or category.

---

## 4. UI Alert Fallback Pattern

**Rule:** Any function callable from the menu must wrap `SpreadsheetApp.getUi().alert()` in a try/catch with a `Logger.log()` fallback.

**Why:** `getUi()` throws when called from the script editor or a trigger context. Without the catch, the function crashes at the last step even though all real work succeeded.

```javascript
function runScan() {
  var results = doTheWork_();

  try {
    SpreadsheetApp.getUi().alert('✅ Done. Found ' + results.count + ' items.');
  } catch (e) {
    Logger.log('Scan complete. Found: ' + results.count);
  }
}
```

**Applies to:** `runScan`, `testRun`, `organizeNow`, `checkStockNow`, and any other menu entry point that shows an alert after doing work.

---

## 5. Deduplication Pattern

**Rule:** Any script that appends rows (log sheets, invoice history, alert logs) must track processed IDs via `PropertiesService` to prevent duplicate entries on repeated runs.

```javascript
var PROP_PROCESSED = 'script_processed_ids';

function getProcessedIds_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PROP_PROCESSED);
  return raw ? JSON.parse(raw) : [];
}

function markProcessed_(ids) {
  var existing = getProcessedIds_();
  var updated = existing.concat(ids).slice(-5000); // cap at 5000
  PropertiesService.getScriptProperties().setProperty(PROP_PROCESSED, JSON.stringify(updated));
}
```

**Applies to:** email-invoice-detector, attachment-auto-saver, follow-up-nudger, vip-priority-alert, smart-ooo-auto-reply.

---

## 6. Brand Constants

**Rule:** Copy this BRAND object into every script. Never hardcode hex values inline.

```javascript
var BRAND = {
  darkBg:      '#1A1A1A',
  gold:        '#C9A84C',
  white:       '#FFFFFF',
  lightGray:   '#F9F9F9',
  successBg:   '#E8F5E9', successText: '#2E7D32',
  warningBg:   '#FFF8E1', warningText: '#F57F17',
  errorBg:     '#FFEBEE', errorText:   '#C62828',
  infoBg:      '#E3F2FD', infoText:    '#1565C0',
  headerFont:  'Roboto Mono',
  bodyFont:    'Roboto',
};
```

---

## 7. Script Version & Header

**Rule:** Every script starts with this header block. Bump version on every published change.

```javascript
/**
 * [Script Name] by TAKScripts
 * ====================================
 * [One line description]
 *
 * Version: 2.0
 * By TAK Ventures — takscripts.store
 */
```

---

## 8. Menu Structure

**Rule:** Every script follows this menu order:

```javascript
ui.createMenu('🕷 TAKScripts')
  .addItem('⚙️ Settings', 'showSettings')
  .addItem('▶️ [Main Action]', 'mainAction')
  .addSeparator()
  .addItem('📊 View Dashboard', 'viewDashboard')
  .addItem('🔄 Refresh Stats', 'refreshDashboardStats')
  .addSeparator()
  .addItem('ℹ️ About TAKScripts', 'showAbout')
  .addToUi();
```

Additional items (Test Run, Undo, etc.) go between Settings/Main Action and the dashboard items.

---

## 9. Settings Load — Always Merge with Defaults

**Rule:** `loadSettings()` must always merge saved values with defaults so new fields added in updates don't break existing installs.

```javascript
function loadSettings() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(SETTINGS_KEY);
  if (!raw) return getDefaultSettings_();
  var saved = JSON.parse(raw);
  var defaults = getDefaultSettings_();
  for (var key in defaults) {
    if (saved[key] === undefined) saved[key] = defaults[key];
  }
  return saved;
}
```

---

## 10. Deliverable File Format

**Rule:** Scripts are delivered to customers as `.txt` files, not `.gs` files.

**Why:** `.txt` files are universally openable. Customers copy the contents and paste into their Apps Script editor.

**Upload path in Supabase Storage:**
```
deliverables/{slug}/{slug}.txt
```

**`product_ideas` fields to set on approval:**
```
deliverable_url  = 'deliverables/{slug}/{slug}.txt'
setup_guide_url  = 'deliverables/{slug}/SETUP.md'   (once guide is written)
stripe_price_id  = (from Stripe dashboard)
is_active        = true
status           = 'live'
```

---

## 11. ES5 vs ES6

**Rule:** Default to ES5 (`var`, `function`, string concatenation) for consistency across the catalog. Only use ES6 if the script was originally written in ES6 — don't mix styles within a file.

Scripts using ES6: `vip-priority-alert`, `data-cleanup-wizard`
All others: ES5
