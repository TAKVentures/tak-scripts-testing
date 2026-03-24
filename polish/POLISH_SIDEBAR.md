# Google Apps Script Sidebar & Modal Dialog — Professional UI Design Reference

**Scope:** Settings panels, help guides, and results displays inside Google Sheets/Docs sidebars and modal dialogs built with HtmlService.
**Audience:** Developer polishing HTML UI across 11+ scripts.
**Format:** Exact CSS values, colors, and patterns throughout — no vague guidance.

---

## 1. Layout Principles

### Hard Constraints

| Surface | Width | Height |
|---|---|---|
| Sidebar | Fixed 300px (Google-enforced, no override) | Full viewport height, scrollable |
| Modal dialog | You control via `setWidth()` / `setHeight()` in `.gs` | Max ~90vw / ~90vh in practice |
| Custom dialog | Same as modal | Same |

The sidebar iframe is exactly 300px wide. Google adds ~6px of internal padding on each side, leaving roughly **288px of usable content width**. Design to **280px usable** to be safe.

### Sidebar Layout System

Use an 8px base grid throughout. All spacing should be multiples of 4 or 8.

```css
/* Base layout — sidebar root */
* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Google Sans', 'Inter', system-ui, -apple-system, sans-serif;
  font-size: 13px;
  line-height: 1.5;
  color: #202124;
  background: #f8f9fa;
  width: 100%;
  overflow-x: hidden;
}

.sidebar-header {
  background: #1a73e8;          /* Google blue — or your brand color */
  color: #ffffff;
  padding: 14px 16px 12px;
  position: sticky;
  top: 0;
  z-index: 100;
}

.sidebar-body {
  padding: 0;                   /* Sections handle their own padding */
  overflow-y: auto;
}
```

### Content Hierarchy Layers

1. **Sticky header** — script name + version or context label. Never taller than 56px.
2. **Section groups** — collapsible or plain dividers, 16px horizontal padding inside.
3. **Form rows** — label + control pairs, 12px vertical gap between rows.
4. **Action bar** — primary button(s) at bottom, sticky-bottom optional.

### Modal Dialog Layout

Modals have no fixed-width constraint. Use these as defaults:

```javascript
// In .gs file
const ui = SpreadsheetApp.getUi();
const html = HtmlService.createHtmlOutputFromFile('MyModal')
  .setWidth(480)
  .setHeight(540);
ui.showModalDialog(html, 'Dialog Title');
```

Inside the modal HTML, center the content:

```css
.modal-container {
  max-width: 440px;
  margin: 0 auto;
  padding: 24px;
}
```

---

## 2. Typography

### Font Stack

```css
/* Primary — Google Sans if available in the Apps Script context, falls back gracefully */
font-family: 'Google Sans', 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;

/* Monospace — for showing script output, JSON, field values */
font-family: 'Roboto Mono', 'Consolas', 'Courier New', monospace;
```

Note: You cannot load Google Fonts via `@import` inside Apps Script sidebars reliably because CSP blocks external font loading in some contexts. Use the system stack above — `system-ui` resolves to San Francisco on Mac, Segoe UI on Windows, Roboto on Chrome OS (where most G Suite users are). This is intentional and looks native.

### Type Scale

| Role | Size | Weight | Color | Line Height |
|---|---|---|---|---|
| Section header | 11px | 600 | #5f6368 | 1.4 |
| Label | 12px | 500 | #3c4043 | 1.4 |
| Input text | 13px | 400 | #202124 | 1.5 |
| Button text | 13px | 500 | varies | 1 |
| Helper/hint text | 11px | 400 | #80868b | 1.4 |
| Error text | 11px | 400 | #d93025 | 1.4 |
| Body/description | 13px | 400 | #3c4043 | 1.6 |
| Results/output | 12px | 400 | #202124 | 1.6 |
| Header title | 14px | 600 | #ffffff | 1.3 |
| Header subtitle | 11px | 400 | rgba(255,255,255,0.8) | 1.3 |

```css
/* Apply the scale */
.label          { font-size: 12px; font-weight: 500; color: #3c4043; }
.helper-text    { font-size: 11px; font-weight: 400; color: #80868b; margin-top: 3px; }
.error-text     { font-size: 11px; font-weight: 400; color: #d93025; margin-top: 3px; }
.section-title  { font-size: 11px; font-weight: 600; color: #5f6368; letter-spacing: 0.5px; text-transform: uppercase; }
```

### What to Avoid

- Font sizes below 11px — illegible in sidebar context
- `font-weight: 700` on anything except errors or emphasis — looks aggressive in small UI
- Italic text for functional UI elements — reserve for quotes or examples only
- All-caps body text — only acceptable for section header labels at 11px with `letter-spacing: 0.5px`

---

## 3. Form Design

### Input Fields

```css
.form-group {
  display: flex;
  flex-direction: column;
  gap: 4px;
  margin-bottom: 14px;
}

.form-group label {
  font-size: 12px;
  font-weight: 500;
  color: #3c4043;
  line-height: 1.4;
}

input[type="text"],
input[type="email"],
input[type="number"],
input[type="url"],
input[type="password"],
textarea,
select {
  width: 100%;
  padding: 7px 10px;
  font-size: 13px;
  font-family: inherit;
  color: #202124;
  background: #ffffff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  outline: none;
  transition: border-color 0.15s ease, box-shadow 0.15s ease;
  appearance: none;
  -webkit-appearance: none;
}

input:hover,
textarea:hover,
select:hover {
  border-color: #b0b8c1;
}

input:focus,
textarea:focus,
select:focus {
  border-color: #1a73e8;
  box-shadow: 0 0 0 3px rgba(26, 115, 232, 0.15);
}

input:disabled,
textarea:disabled,
select:disabled {
  background: #f1f3f4;
  color: #9aa0a6;
  border-color: #e8eaed;
  cursor: not-allowed;
}

input::placeholder,
textarea::placeholder {
  color: #9aa0a6;
  font-weight: 400;
}
```

Key values:
- **Border:** `1px solid #dadce0` (Google's standard gray border)
- **Border radius:** `6px` for inputs; `8px` for cards/sections; `4px` for small badges
- **Padding:** `7px 10px` — matches Google's own Workspace UI targets
- **Focus ring:** `box-shadow: 0 0 0 3px rgba(26,115,232,0.15)` — visible but not harsh

### Textarea

```css
textarea {
  min-height: 72px;
  max-height: 200px;
  resize: vertical;
  line-height: 1.5;
}
```

For results/output textareas:

```css
textarea.output {
  font-family: 'Roboto Mono', monospace;
  font-size: 11px;
  background: #f8f9fa;
  border-color: #e8eaed;
  resize: none;
}
```

### Select / Dropdown

```css
.select-wrapper {
  position: relative;
}

.select-wrapper::after {
  content: '';
  position: absolute;
  right: 10px;
  top: 50%;
  transform: translateY(-50%);
  width: 0;
  height: 0;
  border-left: 4px solid transparent;
  border-right: 4px solid transparent;
  border-top: 5px solid #5f6368;
  pointer-events: none;
}

select {
  padding-right: 28px;  /* room for the arrow */
  cursor: pointer;
}
```

### Checkboxes

```css
.checkbox-group {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  margin-bottom: 10px;
}

.checkbox-group input[type="checkbox"] {
  width: 16px;
  height: 16px;
  min-width: 16px;
  margin-top: 1px;
  accent-color: #1a73e8;
  cursor: pointer;
  border-radius: 3px;
}

.checkbox-group .checkbox-label {
  font-size: 13px;
  color: #3c4043;
  line-height: 1.4;
  cursor: pointer;
  user-select: none;
}

.checkbox-group .checkbox-hint {
  font-size: 11px;
  color: #80868b;
  margin-top: 2px;
}
```

### Toggle Switch (CSS-only)

```css
.toggle-group {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
}

.toggle-label-block { flex: 1; }
.toggle-label-block .label { font-size: 13px; font-weight: 500; color: #202124; }
.toggle-label-block .hint  { font-size: 11px; color: #80868b; margin-top: 2px; }

.toggle {
  position: relative;
  display: inline-block;
  width: 36px;
  height: 20px;
  flex-shrink: 0;
}

.toggle input { opacity: 0; width: 0; height: 0; position: absolute; }

.toggle-track {
  position: absolute;
  inset: 0;
  background: #dadce0;
  border-radius: 20px;
  transition: background 0.2s ease;
  cursor: pointer;
}

.toggle-track::before {
  content: '';
  position: absolute;
  width: 14px;
  height: 14px;
  left: 3px;
  top: 3px;
  background: #ffffff;
  border-radius: 50%;
  transition: transform 0.2s ease;
  box-shadow: 0 1px 3px rgba(0,0,0,0.2);
}

.toggle input:checked + .toggle-track { background: #1a73e8; }
.toggle input:checked + .toggle-track::before { transform: translateX(16px); }
.toggle input:focus-visible + .toggle-track { outline: 2px solid #1a73e8; outline-offset: 2px; }
```

### Field Spacing Rules

- Between label and input: `4px` (gap in flex column)
- Between form groups: `14px`
- Between a section header and first field: `12px`
- Between helper/error text and the next field: included in the `14px` margin-bottom on `.form-group`
- Never use `<br>` for spacing — use margin/padding exclusively

---

## 4. Color Scheme

### Recommended Palette

```css
:root {
  /* Brand / Primary */
  --color-primary:        #1a73e8;
  --color-primary-hover:  #1557b0;
  --color-primary-light:  #e8f0fe;
  --color-primary-focus:  rgba(26, 115, 232, 0.15);

  /* Header */
  --color-header-bg:      #1a73e8;      /* or #1e2a3a for dark navy */
  --color-header-text:    #ffffff;
  --color-header-sub:     rgba(255, 255, 255, 0.75);

  /* Body backgrounds */
  --color-bg-page:        #f8f9fa;      /* sidebar body bg */
  --color-bg-card:        #ffffff;      /* section/card bg */
  --color-bg-subtle:      #f1f3f4;      /* disabled, code bg, output areas */

  /* Text */
  --color-text-primary:   #202124;
  --color-text-secondary: #3c4043;
  --color-text-muted:     #5f6368;
  --color-text-placeholder: #9aa0a6;

  /* Borders */
  --color-border:         #dadce0;
  --color-border-hover:   #b0b8c1;
  --color-border-focus:   #1a73e8;
  --color-border-section: #e8eaed;

  /* Status */
  --color-success:        #137333;
  --color-success-bg:     #e6f4ea;
  --color-success-border: #a8d5b5;
  --color-error:          #d93025;
  --color-error-bg:       #fce8e6;
  --color-error-border:   #f5c6c3;
  --color-warning:        #b06000;
  --color-warning-bg:     #fef7e0;
  --color-warning-border: #f6d89a;
  --color-info:           #1a73e8;
  --color-info-bg:        #e8f0fe;
  --color-info-border:    #aecbfa;
}
```

### Dark Navy Header Alternative

```css
--color-header-bg:   #1e2a3a;   /* dark navy — more premium, less "Google" */
--color-header-text: #ffffff;
--color-header-sub:  rgba(255, 255, 255, 0.65);
```

---

## 5. Button Design

### Full Button System

```css
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 8px 16px;
  font-size: 13px;
  font-weight: 500;
  font-family: inherit;
  line-height: 1;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.15s ease, box-shadow 0.15s ease, opacity 0.15s ease;
  white-space: nowrap;
  user-select: none;
}

.btn-block { width: 100%; }

/* Primary */
.btn-primary { background: #1a73e8; color: #ffffff; }
.btn-primary:hover:not(:disabled) { background: #1557b0; box-shadow: 0 1px 4px rgba(26,115,232,0.35); }
.btn-primary:active:not(:disabled) { background: #0d47a1; box-shadow: none; }

/* Secondary */
.btn-secondary { background: #ffffff; color: #1a73e8; border: 1px solid #dadce0; }
.btn-secondary:hover:not(:disabled) { background: #f0f4ff; border-color: #aecbfa; }

/* Danger */
.btn-danger { background: #d93025; color: #ffffff; }
.btn-danger:hover:not(:disabled) { background: #b31c12; box-shadow: 0 1px 4px rgba(217,48,37,0.35); }

/* Ghost */
.btn-ghost { background: transparent; color: #1a73e8; padding: 8px 8px; }
.btn-ghost:hover:not(:disabled) { background: #f0f4ff; }

/* Disabled */
.btn:disabled { opacity: 0.45; cursor: not-allowed; box-shadow: none; }

/* Focus */
.btn:focus-visible { outline: 2px solid #1a73e8; outline-offset: 2px; }
```

### Loading State (BUILD STANDARD — Rule 13)

```css
.btn-loading {
  position: relative;
  color: transparent !important;
  pointer-events: none;
}

.btn-loading::after {
  content: '';
  position: absolute;
  width: 14px;
  height: 14px;
  top: 50%; left: 50%;
  margin-top: -7px; margin-left: -7px;
  border: 2px solid rgba(255,255,255,0.4);
  border-top-color: #ffffff;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
}

@keyframes spin { to { transform: rotate(360deg); } }
```

**Save Confirmation Pattern (Rule 13):**
```javascript
function handleSave() {
  const btn = document.getElementById('saveBtn');
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Saving…';

  google.script.run
    .withSuccessHandler(function() {
      btn.textContent = '✓ Saved!';
      setTimeout(function() {
        btn.disabled = false;
        btn.textContent = originalText;
      }, 2500);
    })
    .withFailureHandler(function(err) {
      btn.disabled = false;
      btn.textContent = originalText;
      showStatus('error', err.message || 'Save failed.');
    })
    .saveSettings(getFormValues());
}
```

### Button Sizing Reference

| Size | Padding | Font | Use case |
|---|---|---|---|
| Small | `5px 10px` | 12px | Inline actions, badge buttons |
| Default | `8px 16px` | 13px | Standard sidebar actions |
| Large | `10px 20px` | 14px | Primary CTA in modal dialogs |

### Sidebar Button Rules

- Primary action: full-width, bottom of the panel or section
- Multiple actions: primary full-width on top, secondary ghost/outline below — never side by side in 280px
- Destructive: always separated from primary; consider confirmation step

---

## 6. Status & Feedback UI

### Status Banner

```css
.status-banner {
  display: none;
  align-items: flex-start;
  gap: 8px;
  padding: 10px 12px;
  border-radius: 6px;
  font-size: 12px;
  line-height: 1.5;
  margin-bottom: 12px;
  border: 1px solid;
}

.status-banner.visible { display: flex; }
.status-banner.success { background: #e6f4ea; border-color: #a8d5b5; color: #137333; }
.status-banner.error   { background: #fce8e6; border-color: #f5c6c3; color: #c5221f; }
.status-banner.warning { background: #fef7e0; border-color: #f6d89a; color: #b06000; }
.status-banner.info    { background: #e8f0fe; border-color: #aecbfa; color: #1558d6; }
```

```javascript
function showStatus(type, message) {
  const banner = document.getElementById('statusBanner');
  const icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };
  banner.className = 'status-banner visible ' + type;
  banner.innerHTML = '<span>' + icons[type] + '</span><span>' + message + '</span>';
  if (type === 'success') {
    setTimeout(function() { banner.classList.remove('visible'); }, 4000);
  }
}
```

### Loading Overlay

```css
.loading-overlay {
  display: none;
  position: fixed;
  inset: 0;
  background: rgba(255, 255, 255, 0.8);
  z-index: 200;
  align-items: center;
  justify-content: center;
  flex-direction: column;
  gap: 12px;
}
.loading-overlay.visible { display: flex; }
.loading-spinner {
  width: 32px; height: 32px;
  border: 3px solid #e8eaed;
  border-top-color: #1a73e8;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}
.loading-text { font-size: 13px; color: #5f6368; font-weight: 500; }
```

### Inline Field Validation

```css
.form-group.has-error input,
.form-group.has-error textarea,
.form-group.has-error select {
  border-color: #d93025;
  box-shadow: 0 0 0 3px rgba(217, 48, 37, 0.1);
}

.form-group.has-error .error-text { display: block; }
.error-text { display: none; font-size: 11px; color: #d93025; margin-top: 3px; }
```

---

## 7. Section Headers & Visual Grouping

```css
.settings-section {
  background: #ffffff;
  border-bottom: 1px solid #e8eaed;
  padding: 14px 16px;
}
.settings-section:last-child { border-bottom: none; }

.section-title {
  font-size: 11px;
  font-weight: 600;
  color: #5f6368;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-bottom: 12px;
}
```

**Section spacing guidelines:**
- Sidebar: sections stack flush with `border-bottom` dividers
- Section internal padding: `14px 16px`
- First section after sticky header: no special treatment needed

---

## 8. Sticky Action Bar

```css
.action-bar {
  position: sticky;
  bottom: 0;
  background: #ffffff;
  border-top: 1px solid #e8eaed;
  padding: 12px 16px;
  z-index: 50;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
```

This keeps the primary action always visible regardless of scroll position — critical for tall settings panels.

---

## 9. Icons & Emoji in UI

### Recommended Unicode Symbols

| Purpose | Symbol | Code |
|---|---|---|
| Success | ✓ | `&#10003;` |
| Error / close | ✕ | `&#10005;` |
| Warning | ⚠ | `&#9888;` |
| Info | ℹ | `&#8505;` |
| Settings | ⚙ | `&#9881;` |
| Arrow right | → | `&#8594;` |
| Collapse | ▾ | `&#9662;` |
| Expand | ▸ | `&#9658;` |

### When Emoji Work
- Section headers: 1 emoji max (e.g., `⚙ General Settings`)
- Button labels: 1 emoji max (e.g., `▶ Run`)
- Empty states: 1 emoji (e.g., `📭 No results yet`)

### When Emoji Hurt
- Error messages (undermines seriousness)
- Form labels (looks unprofessional)
- More than 3 section headers in a row
- As sole status indicator (renders inconsistently cross-OS)

---

## 10. Scrollable Areas

```css
/* Constrain growing content sections */
.results-list {
  max-height: 240px;
  overflow-y: auto;
  border: 1px solid #e8eaed;
  border-radius: 6px;
  background: #f8f9fa;
}

/* Styled scrollbar */
.results-list::-webkit-scrollbar { width: 6px; }
.results-list::-webkit-scrollbar-track { background: transparent; }
.results-list::-webkit-scrollbar-thumb { background: #dadce0; border-radius: 3px; }
.results-list::-webkit-scrollbar-thumb:hover { background: #b0b8c1; }
```

**Rules:**
- Never allow horizontal scroll — design single-column only
- Overflow content areas get `max-height` + `overflow-y: auto`
- Body never clips — let it scroll freely; only constrain sub-panels

---

## 11. Accessibility

### Contrast Ratios (WCAG AA)

| Element | Foreground | Background | Ratio |
|---|---|---|---|
| Body text | `#202124` | `#ffffff` | 16.1:1 ✓ |
| Secondary text | `#3c4043` | `#ffffff` | 12.6:1 ✓ |
| Muted text | `#5f6368` | `#ffffff` | 7.0:1 ✓ |
| Error text | `#d93025` | `#fce8e6` | 4.8:1 ✓ |
| Success text | `#137333` | `#e6f4ea` | 5.9:1 ✓ |
| White on primary | `#ffffff` | `#1a73e8` | 3.9:1 ✓ |

### Key Rules
- Every `<input>`, `<select>`, `<textarea>` has an explicit `<label>` with matching `for`/`id`
- Never use placeholder as label substitute
- Error messages linked via `aria-describedby`
- Color never the sole status indicator — always pair with text or icon
- Status banners use `role="alert"` and `aria-live="polite"`

```css
:focus-visible { outline: 2px solid #1a73e8; outline-offset: 2px; }
:focus:not(:focus-visible) { outline: none; }
```

---

## 12. What Makes Sidebar UI Feel Cheap vs. Premium

### What Makes It Feel Cheap

| Anti-pattern | Fix |
|---|---|
| Default `<select>` styling | `appearance: none` + custom arrow |
| `<hr>` with browser default styles | `border: none; height: 1px; background: #e8eaed` |
| `alert()` / `confirm()` for feedback | In-panel status banners |
| Button text in ALL CAPS | `text-transform: none; font-weight: 500` |
| Inconsistent spacing (10px here, 8px there) | 8px base grid only |
| Red `#ff0000` for errors | `#d93025` |
| No loading state on async actions | Spinner + button disabled |
| `outline: none` globally | `:focus-visible` ring always present |
| `<table>` for layout | Flexbox only |
| No disabled state during processing | Always `btn.disabled = true` before async |
| Placeholder as label substitute | Always a separate visible `<label>` |
| Heavy shadows (`box-shadow: 0 4px 20px ...`) | Max `0 1px 4px rgba(0,0,0,0.12)` |
| Emoji in every section header | Max 1 emoji per section |
| Pure `#ffffff` sidebar body | `#f8f9fa` body + `#ffffff` cards |
| `font-weight: bold` everywhere | 400 body, 500 labels, 600 section headers |

### What Makes It Feel Premium

- **8px grid rhythm** — every element sits on the same invisible baseline
- **Micro-transitions** — `transition: border-color 0.15s, box-shadow 0.15s` on inputs
- **Smart helper text** — one concise line only where truly needed
- **Sticky header + sticky action bar** — structure stays, content scrolls
- **Loading state showing progress** — "Running, please wait…" beats silent disable
- **Success state auto-dismisses** — `✓ Saved` for 4s, then fades out
- **Error messages that say what to do** — not "Error: 403" but "API key invalid. Check Settings → API Key."

---

## 13. Complete Starter Template

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --primary: #1a73e8;
      --primary-hover: #1557b0;
      --primary-focus: rgba(26,115,232,0.15);
      --text-primary: #202124;
      --text-secondary: #3c4043;
      --text-muted: #5f6368;
      --text-placeholder: #9aa0a6;
      --border: #dadce0;
      --border-section: #e8eaed;
      --bg-page: #f8f9fa;
      --bg-card: #ffffff;
      --bg-disabled: #f1f3f4;
    }

    body {
      font-family: 'Google Sans', 'Segoe UI', system-ui, sans-serif;
      font-size: 13px; line-height: 1.5;
      color: var(--text-primary);
      background: var(--bg-page);
      width: 100%; overflow-x: hidden;
    }

    .sidebar-header {
      background: var(--primary); color: #fff;
      padding: 12px 16px;
      position: sticky; top: 0; z-index: 100;
    }
    .sidebar-header h1 { font-size: 14px; font-weight: 600; line-height: 1.3; }
    .sidebar-header .subtitle { font-size: 11px; opacity: 0.75; margin-top: 1px; }

    .status-banner {
      display: none; align-items: flex-start; gap: 8px;
      padding: 10px 12px; border-radius: 6px; font-size: 12px;
      line-height: 1.5; border: 1px solid; margin: 12px 16px 0;
    }
    .status-banner.visible { display: flex; }
    .status-banner.success { background: #e6f4ea; border-color: #a8d5b5; color: #137333; }
    .status-banner.error   { background: #fce8e6; border-color: #f5c6c3; color: #c5221f; }
    .status-banner.warning { background: #fef7e0; border-color: #f6d89a; color: #b06000; }
    .status-banner.info    { background: #e8f0fe; border-color: #aecbfa; color: #1558d6; }

    .settings-section {
      background: var(--bg-card);
      border-bottom: 1px solid var(--border-section);
      padding: 14px 16px;
    }
    .settings-section:last-of-type { border-bottom: none; }
    .section-title {
      font-size: 11px; font-weight: 600; color: var(--text-muted);
      text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px;
    }

    .form-group { display: flex; flex-direction: column; gap: 4px; margin-bottom: 14px; }
    .form-group:last-child { margin-bottom: 0; }
    label { font-size: 12px; font-weight: 500; color: var(--text-secondary); }

    input[type="text"], input[type="email"], input[type="number"],
    input[type="url"], input[type="password"], textarea, select {
      width: 100%; padding: 7px 10px; font-size: 13px; font-family: inherit;
      color: var(--text-primary); background: var(--bg-card);
      border: 1px solid var(--border); border-radius: 6px; outline: none;
      transition: border-color 0.15s, box-shadow 0.15s;
      appearance: none; -webkit-appearance: none;
    }
    input:focus, textarea:focus, select:focus {
      border-color: var(--primary);
      box-shadow: 0 0 0 3px var(--primary-focus);
    }
    input:disabled, textarea:disabled, select:disabled {
      background: var(--bg-disabled); color: #9aa0a6;
      border-color: #e8eaed; cursor: not-allowed;
    }
    input::placeholder, textarea::placeholder { color: var(--text-placeholder); }
    .helper-text { font-size: 11px; color: var(--text-placeholder); margin-top: 3px; }
    .error-text { display: none; font-size: 11px; color: #d93025; margin-top: 3px; }
    .form-group.has-error input, .form-group.has-error select {
      border-color: #d93025; box-shadow: 0 0 0 3px rgba(217,48,37,0.1);
    }
    .form-group.has-error .error-text { display: block; }
    textarea { min-height: 72px; resize: vertical; }

    .btn {
      display: inline-flex; align-items: center; justify-content: center;
      gap: 6px; padding: 8px 16px; font-size: 13px; font-weight: 500;
      font-family: inherit; line-height: 1; border: none; border-radius: 6px;
      cursor: pointer; transition: background 0.15s, box-shadow 0.15s;
      white-space: nowrap; user-select: none;
    }
    .btn-block { width: 100%; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-primary:hover:not(:disabled) { background: var(--primary-hover); box-shadow: 0 1px 4px rgba(26,115,232,0.35); }
    .btn-secondary { background: #fff; color: var(--primary); border: 1px solid var(--border); }
    .btn-secondary:hover:not(:disabled) { background: #f0f4ff; border-color: #aecbfa; }
    .btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .btn:focus-visible { outline: 2px solid var(--primary); outline-offset: 2px; }
    .btn-loading { position: relative; color: transparent !important; pointer-events: none; }
    .btn-loading::after {
      content: ''; position: absolute; width: 14px; height: 14px;
      top: 50%; left: 50%; margin: -7px 0 0 -7px;
      border: 2px solid rgba(255,255,255,0.4); border-top-color: #fff;
      border-radius: 50%; animation: spin 0.7s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .action-bar {
      position: sticky; bottom: 0; background: #fff;
      border-top: 1px solid var(--border-section);
      padding: 12px 16px; z-index: 50;
      display: flex; flex-direction: column; gap: 8px;
    }

    :focus-visible { outline: 2px solid var(--primary); outline-offset: 2px; }
    :focus:not(:focus-visible) { outline: none; }
  </style>
</head>
<body>

  <div class="sidebar-header">
    <h1>Script Name</h1>
    <div class="subtitle">Settings · v1.0</div>
  </div>

  <div id="statusBanner" class="status-banner" role="alert" aria-live="polite"></div>

  <div class="settings-section">
    <div class="section-title">⚙ General</div>

    <div class="form-group">
      <label for="inputField">Field Label</label>
      <input type="text" id="inputField" placeholder="Enter value…">
      <span class="helper-text">Short explanation of what this does.</span>
      <span class="error-text">This field is required.</span>
    </div>

    <div class="form-group">
      <label for="mySelect">Select Option</label>
      <select id="mySelect">
        <option value="">Choose one…</option>
        <option value="a">Option A</option>
        <option value="b">Option B</option>
      </select>
    </div>
  </div>

  <div class="action-bar">
    <button class="btn btn-primary btn-block" id="saveBtn" onclick="handleSave()">Save Settings</button>
  </div>

  <script>
    function showStatus(type, message) {
      var b = document.getElementById('statusBanner');
      var icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };
      b.className = 'status-banner visible ' + type;
      b.innerHTML = '<span>' + icons[type] + '</span><span>' + message + '</span>';
      if (type === 'success') {
        setTimeout(function() { b.classList.remove('visible'); }, 4000);
      }
    }

    function handleSave() {
      var btn = document.getElementById('saveBtn');
      var original = btn.textContent;
      btn.disabled = true;
      btn.textContent = 'Saving…';

      google.script.run
        .withSuccessHandler(function() {
          btn.textContent = '✓ Saved!';
          setTimeout(function() {
            btn.disabled = false;
            btn.textContent = original;
          }, 2500);
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.textContent = original;
          showStatus('error', err.message || 'Save failed. Please try again.');
        })
        .saveSettings();
    }
  </script>
</body>
</html>
```

---

## 14. Modal Dialog Specifics

```css
body {
  font-family: 'Google Sans', 'Segoe UI', system-ui, sans-serif;
  font-size: 13px; color: #202124;
  background: #ffffff;          /* Modals = white, not #f8f9fa */
  padding: 0; margin: 0;
}

.modal-container { padding: 24px; max-width: 440px; margin: 0 auto; }
.modal-title { font-size: 16px; font-weight: 600; color: #202124; margin-bottom: 6px; }
.modal-subtitle { font-size: 13px; color: #5f6368; margin-bottom: 20px; line-height: 1.5; }
.modal-footer {
  display: flex; justify-content: flex-end; gap: 8px;
  margin-top: 20px; padding-top: 16px;
  border-top: 1px solid #e8eaed;
}
```

**For help guide modals:**

```css
.help-section { margin-bottom: 20px; }
.help-section h3 { font-size: 13px; font-weight: 600; color: #202124; margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
.help-section p, .help-section li { font-size: 13px; color: #3c4043; line-height: 1.6; }
.help-section ol, .help-section ul { padding-left: 18px; margin-top: 6px; }
.help-section li { margin-bottom: 4px; }
.help-note {
  background: #e8f0fe; border-left: 3px solid #1a73e8;
  padding: 10px 12px; border-radius: 0 6px 6px 0;
  font-size: 12px; color: #1558d6; margin-top: 12px;
}
```

---

## Quick Reference Card

| Property | Value |
|---|---|
| Sidebar usable width | ~280px |
| Body background | `#f8f9fa` |
| Card/section background | `#ffffff` |
| Base font | `'Google Sans', 'Segoe UI', system-ui, sans-serif` |
| Base font size | `13px` |
| Label size | `12px / weight 500` |
| Section header | `11px / weight 600 / uppercase / #5f6368` |
| Helper text | `11px / #80868b` |
| Error text | `11px / #d93025` |
| Input border | `1px solid #dadce0` |
| Input border-radius | `6px` |
| Input padding | `7px 10px` |
| Focus ring | `box-shadow: 0 0 0 3px rgba(26,115,232,0.15)` |
| Base grid unit | `8px` |
| Section internal padding | `14px 16px` |
| Form group spacing | `14px margin-bottom` |
| Button border-radius | `6px` |
| Button font | `13px / weight 500` |
| Primary button | `bg #1a73e8 / hover #1557b0` |
| Divider color | `#e8eaed` |
| Success green | `#137333 / bg #e6f4ea` |
| Error red | `#d93025 / bg #fce8e6` |
| Warning amber | `#b06000 / bg #fef7e0` |
| Info blue | `#1558d6 / bg #e8f0fe` |
