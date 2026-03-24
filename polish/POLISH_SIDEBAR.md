# POLISH_SIDEBAR — Google Apps Script Sidebar & Dialog Design Reference

A comprehensive design reference for building polished, professional HTML sidebars and modal dialogs inside Google Sheets and Docs via `HtmlService`. All values are specific and implementation-ready.

---

## Table of Contents

1. [Sidebar & Dialog Dimensions](#1-sidebar--dialog-dimensions)
2. [CSS Reset & Base Styles](#2-css-reset--base-styles)
3. [Color Palette](#3-color-palette)
4. [Layout Structure](#4-layout-structure)
5. [Typography](#5-typography)
6. [Form Elements](#6-form-elements)
7. [Button Styles](#7-button-styles)
8. [Save Confirmation Pattern](#8-save-confirmation-pattern)
9. [Status Indicators & Badges](#9-status-indicators--badges)
10. [Tables in Sidebars](#10-tables-in-sidebars)
11. [Loading States](#11-loading-states)
12. [Error & Success Messages](#12-error--success-messages)
13. [Help Sections & Tooltips](#13-help-sections--tooltips)
14. [Scrolling & Sticky Layout](#14-scrolling--sticky-layout)
15. [Google Apps Script Integration](#15-google-apps-script-integration)
16. [Full Sample Sidebar Template](#16-full-sample-sidebar-template)
17. [What to Avoid](#17-what-to-avoid)

---

## 1. Sidebar & Dialog Dimensions

### Standard Sidebar

The default sidebar opened with `SpreadsheetApp.getUi().showSidebar()` is **fixed at 300px wide** by the Google Sheets container. You cannot change this width — Google enforces it regardless of what you pass to `setWidth()`. Design for 300px as the absolute ceiling.

- **Width:** 300px (fixed by Google — do not try to override)
- **Height:** 100% of the browser viewport minus Google's own chrome
- **Usable content width:** 268px after applying 16px padding on each side
- **Rule:** All form elements must be `width: 100%` with `box-sizing: border-box` so they fill the usable width without overflowing

The sidebar renders inside a sandboxed iframe. Parent page CSS does not bleed in. External resources (Google Fonts, CDN assets) load normally as long as the URL is public.

### Wide Sidebar (Custom Dialog Used as Sidebar)

When 300px is too narrow for complex layouts, open a modeless dialog positioned to the side. This gives you width control.

- **Wide sidebar width:** 430px
- **Minimum usable width:** 320px (below this, form fields become cramped)
- **Use case:** Multi-column forms, data-heavy configuration panels, side-by-side comparisons

```javascript
// Wide sidebar via modeless dialog
function showWideSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(430)
    .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Configuration');
}
```

### Modal Dialog Sizes

Modal dialogs opened with `showModalDialog()` block the underlying sheet. Use them for confirmations, critical inputs, and short wizards. Google does not auto-resize modals — always set an explicit height.

| Size | Width | Height | Use Case |
|------|-------|--------|----------|
| Small | 400px | 200px | Single confirmations, one-field prompts |
| Medium | 520px | 380px | Multi-field forms, settings panels |
| Large | 720px | 520px | Data review tables, multi-step wizards |

```javascript
// Small confirmation dialog
function showConfirmDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfirmDialog')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Confirm Action');
}

// Medium settings dialog
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(520)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}
```

### showSidebar vs showModalDialog vs showModelessDialog

| Method | Blocks Sheet? | User Can Dismiss? | Width Control? |
|--------|:------------:|:-----------------:|:--------------:|
| `showSidebar()` | No | Yes (X button) | No (fixed 300px) |
| `showModalDialog()` | Yes | No (must close via code or X) | Yes |
| `showModelessDialog()` | No | Yes (X button) | Yes |

Use `showSidebar()` for persistent tools and configuration panels. Use `showModalDialog()` for decisions that must be answered before proceeding. Use `showModelessDialog()` when you need width control but do not want to block the sheet.

---

## 2. CSS Reset & Base Styles

Apply this reset at the top of every sidebar and dialog `<style>` block. It normalizes browser quirks and establishes consistent box-sizing across all elements.

```css
/* ── Reset & Base ─────────────────────────────────── */
*, *::before, *::after {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

html, body {
  height: 100%;
  overflow: hidden; /* body itself does not scroll — .sidebar-body div does */
}

body {
  font-family: 'Google Sans', Roboto, system-ui, -apple-system, sans-serif;
  font-size: 13px;
  line-height: 1.5;
  color: #202124;
  background: #ffffff;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

a {
  color: #1a73e8;
  text-decoration: none;
}

a:hover {
  text-decoration: underline;
}
```

### CSS Custom Properties (Variables)

Define all design tokens as CSS custom properties at the top of the `<style>` block. This keeps the entire sidebar theme changeable from one place and prevents hard-coded color values from scattering through the stylesheet.

```css
:root {
  /* Brand */
  --color-primary:         #1a73e8;
  --color-primary-dark:    #1557b0;
  --color-primary-light:   #e8f0fe;

  /* Status */
  --color-success:         #137333;
  --color-success-bg:      #e6f4ea;
  --color-warning:         #b06000;
  --color-warning-bg:      #fef7e0;
  --color-error:           #c5221f;
  --color-error-bg:        #fce8e6;
  --color-info:            #1967d2;
  --color-info-bg:         #e8f0fe;

  /* Neutrals */
  --color-text:            #202124;
  --color-text-secondary:  #5f6368;
  --color-text-disabled:   #9aa0a6;
  --color-border:          #dadce0;
  --color-border-focus:    #1a73e8;
  --color-bg-hover:        #f8f9fa;
  --color-bg-input:        #ffffff;

  /* Layout */
  --sidebar-padding:       16px;
  --border-radius:         6px;
  --border-radius-sm:      4px;
}
```

---

## 3. Color Palette

### Primary Colors

| Role | Hex | CSS Variable |
|------|-----|-------------|
| Primary blue | `#1a73e8` | `--color-primary` |
| Primary blue hover/pressed | `#1557b0` | `--color-primary-dark` |
| Primary blue light background | `#e8f0fe` | `--color-primary-light` |

### Status / Semantic Colors

| Status | Text Color | Background | Use Case |
|--------|-----------|-----------|---------|
| Success | `#137333` | `#e6f4ea` | Completed, active, paid, saved |
| Warning | `#b06000` | `#fef7e0` | Pending, at risk, near threshold |
| Error | `#c5221f` | `#fce8e6` | Failed, overdue, invalid input |
| Info | `#1967d2` | `#e8f0fe` | Informational, in progress, neutral |

### Neutral Grays

| Role | Hex |
|------|-----|
| Primary text | `#202124` |
| Secondary text / labels | `#5f6368` |
| Placeholder / disabled text | `#9aa0a6` |
| Border (default) | `#dadce0` |
| Border (subtle hover) | `#bdc1c6` |
| Page / section background | `#f8f9fa` |
| White (inputs, cards) | `#ffffff` |

**Rule:** Never use raw color values like `color: red` or `color: green` in the stylesheet. Always reference a CSS variable or one of the specific hex values above. This makes global palette changes a one-line edit.

---

## 4. Layout Structure

### The Three-Zone Sidebar

Every sidebar uses the same three-zone structure: a sticky header, a scrollable body, and a sticky footer. The body takes all remaining vertical space and scrolls independently. The header and footer are always visible.

```html
<div class="sidebar">

  <!-- Zone 1: Sticky header — title, subtitle, optional status -->
  <div class="sidebar-header">
    <h1 class="sidebar-title">Sidebar Title</h1>
    <p class="sidebar-subtitle">Short description of what this panel does</p>
  </div>

  <!-- Zone 2: Scrollable body — all sections and form content -->
  <div class="sidebar-body">
    <!-- sections and fields go here -->
  </div>

  <!-- Zone 3: Sticky footer — primary action buttons -->
  <div class="sidebar-footer">
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn btn-primary" id="saveBtn" onclick="handleSave()">Save</button>
  </div>

</div>
```

```css
/* ── Three-Zone Layout ────────────────────────────── */
.sidebar {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}

.sidebar-header {
  flex-shrink: 0;
  padding: 16px var(--sidebar-padding) 12px;
  border-bottom: 1px solid var(--color-border);
  background: #ffffff;
}

.sidebar-body {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  padding: 16px var(--sidebar-padding);
  scroll-behavior: smooth;
}

.sidebar-footer {
  flex-shrink: 0;
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 12px var(--sidebar-padding);
  border-top: 1px solid var(--color-border);
  background: #ffffff;
}
```

### Section Blocks

Use `.section` divs inside `.sidebar-body` to group related fields visually. Give each section a `.section-header` label so users understand what each group of fields is for.

```css
.section {
  margin-bottom: 20px;
}

.section:last-child {
  margin-bottom: 0;
}

.section-header {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.5px;
  text-transform: uppercase;
  color: var(--color-text-secondary);
  margin-bottom: 10px;
  padding-bottom: 6px;
  border-bottom: 1px solid var(--color-border);
}
```

```html
<div class="section">
  <div class="section-header">Client Details</div>
  <!-- field groups go here -->
</div>

<div class="section">
  <div class="section-header">Output Settings</div>
  <!-- field groups go here -->
</div>
```

---

## 5. Typography

### Font Sizes

| Element | Size | Weight | Color |
|---------|------|--------|-------|
| Sidebar title | 16px | 500 | `#202124` |
| Sidebar subtitle | 12px | 400 | `#5f6368` |
| Section header | 11px | 600 | `#5f6368` |
| Field label | 12px | 500 | `#3c4043` |
| Input / body text | 13px | 400 | `#202124` |
| Help text | 11px | 400 | `#5f6368` |
| Error message | 11px | 400 | `#c5221f` |
| Button text | 13px | 500 | varies by type |
| Badge / tag text | 11px | 500 | varies by status |

```css
/* ── Typography ───────────────────────────────────── */
.sidebar-title {
  font-size: 16px;
  font-weight: 500;
  color: var(--color-text);
  line-height: 1.3;
}

.sidebar-subtitle {
  font-size: 12px;
  color: var(--color-text-secondary);
  margin-top: 2px;
  line-height: 1.4;
}

.field-label {
  display: block;
  font-size: 12px;
  font-weight: 500;
  color: #3c4043;
  margin-bottom: 4px;
}

.help-text {
  font-size: 11px;
  color: var(--color-text-secondary);
  margin-top: 4px;
  line-height: 1.4;
}

.error-text {
  font-size: 11px;
  color: var(--color-error);
  margin-top: 4px;
  display: none; /* shown via JS when validation fails */
}
```

**Why these sizes:** 13px is the baseline — it matches Google Sheets' own interface text. Labels at 12px are clearly secondary but readable. Section headers at 11px uppercase are meta-labels, not content, so they can be smaller. Help text at 11px signals subordinate information without straining readability at 1x zoom.

---

## 6. Form Elements

### Field Group Wrapper

Wrap every label + input + help-text trio in a `.field-group` div. This keeps vertical spacing consistent across all fields without managing margins on individual elements.

```css
.field-group {
  margin-bottom: 14px;
}

.field-group:last-child {
  margin-bottom: 0;
}
```

### Text Inputs & Selects

```css
input[type="text"],
input[type="email"],
input[type="number"],
input[type="date"],
input[type="url"],
input[type="password"],
select,
textarea {
  width: 100%;
  padding: 7px 10px;
  font-size: 13px;
  font-family: inherit;
  color: var(--color-text);
  background: var(--color-bg-input);
  border: 1px solid var(--color-border);
  border-radius: var(--border-radius-sm);
  outline: none;
  transition: border-color 0.15s ease, box-shadow 0.15s ease;
  appearance: none;
  -webkit-appearance: none;
}

input:focus,
select:focus,
textarea:focus {
  border-color: var(--color-border-focus);
  box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.15);
}

input::placeholder,
textarea::placeholder {
  color: var(--color-text-disabled);
}

input:disabled,
select:disabled,
textarea:disabled {
  background: #f1f3f4;
  color: var(--color-text-disabled);
  cursor: not-allowed;
}
```

**Critical:** `font-family: inherit` must be set on every input and textarea. Chrome's default font inside the sidebar iframe differs from the body font — without this, inputs render in a different typeface than the rest of the sidebar.

### Textarea

```css
textarea {
  resize: vertical;      /* user can expand height but not break horizontal layout */
  min-height: 72px;
  line-height: 1.5;
}
```

### Select Dropdown

`appearance: none` removes the browser's native arrow. Add a custom chevron using an `::after` pseudo-element on a wrapper div.

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
  border-top: 5px solid var(--color-text-secondary);
  pointer-events: none;
}

.select-wrapper select {
  padding-right: 28px; /* room for chevron */
  cursor: pointer;
}
```

```html
<div class="field-group">
  <label class="field-label" for="statusSelect">Status</label>
  <div class="select-wrapper">
    <select id="statusSelect">
      <option value="">— Select —</option>
      <option value="active">Active</option>
      <option value="pending">Pending</option>
      <option value="closed">Closed</option>
    </select>
  </div>
</div>
```

### Checkboxes & Radio Buttons

Use a flex row wrapper to align the control with its label. Set `accent-color` to match the brand so the checked state uses the primary blue.

```css
.checkbox-group,
.radio-group {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.checkbox-item,
.radio-item {
  display: flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
}

.checkbox-item input[type="checkbox"],
.radio-item input[type="radio"] {
  width: 14px;
  height: 14px;
  flex-shrink: 0;
  cursor: pointer;
  accent-color: var(--color-primary);
  margin: 0;
}

.checkbox-item label,
.radio-item label {
  font-size: 13px;
  color: var(--color-text);
  cursor: pointer;
  user-select: none;
}
```

```html
<div class="field-group">
  <div class="checkbox-group">
    <div class="checkbox-item">
      <input type="checkbox" id="optEmail" name="opts" value="email">
      <label for="optEmail">Send email notification</label>
    </div>
    <div class="checkbox-item">
      <input type="checkbox" id="optDrive" name="opts" value="drive">
      <label for="optDrive">Create Drive folder</label>
    </div>
  </div>
</div>
```

---

## 7. Button Styles

### Base Button

All buttons share a base `.btn` class. Variant classes layer on top.

```css
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 7px 16px;
  font-size: 13px;
  font-weight: 500;
  font-family: inherit;
  border-radius: var(--border-radius-sm);
  border: 1px solid transparent;
  cursor: pointer;
  white-space: nowrap;
  transition: background 0.15s ease, box-shadow 0.15s ease, opacity 0.15s ease;
  user-select: none;
  min-width: 72px;
  line-height: 1;
}

.btn:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: 2px;
}

.btn:disabled {
  opacity: 0.55;
  cursor: not-allowed;
  pointer-events: none;
}
```

### Primary Button (Blue)

```css
.btn-primary {
  background: var(--color-primary);
  color: #ffffff;
  border-color: var(--color-primary);
}

.btn-primary:hover:not(:disabled) {
  background: var(--color-primary-dark);
  border-color: var(--color-primary-dark);
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
}

.btn-primary:active:not(:disabled) {
  background: #0d47a1;
  box-shadow: none;
}
```

### Secondary Button (Outlined)

```css
.btn-secondary {
  background: #ffffff;
  color: var(--color-primary);
  border-color: var(--color-border);
}

.btn-secondary:hover:not(:disabled) {
  background: var(--color-primary-light);
  border-color: var(--color-primary);
}
```

### Danger Button (Red)

```css
.btn-danger {
  background: #ffffff;
  color: var(--color-error);
  border-color: #f28b82;
}

.btn-danger:hover:not(:disabled) {
  background: var(--color-error-bg);
  border-color: var(--color-error);
}
```

### Full-Width Button

For single-action footers where the button should span the full footer width:

```css
.btn-full {
  width: 100%;
}
```

---

## 8. Save Confirmation Pattern

This is a **BUILD_STANDARD** for all TAKScripts sidebars. Every save action must follow this exact sequence — no exceptions.

**Sequence:**
1. User clicks Save
2. Button is immediately disabled and label changes to "Saving…"
3. On success: label changes to "✓ Saved!" and button turns green
4. After 2500ms: button returns to normal enabled state
5. On failure: button re-enables, label restores to "Save", error banner appears

```html
<button class="btn btn-primary" id="saveBtn" onclick="handleSave()">Save</button>
```

```javascript
function handleSave() {
  if (!validateForm()) return;

  const btn = document.getElementById('saveBtn');
  clearAlerts();
  setSavingState(btn);

  const data = {
    field1: document.getElementById('field1').value.trim(),
    field2: document.getElementById('field2').value.trim()
  };

  google.script.run
    .withUserObject(btn)
    .withSuccessHandler(function(result, btn) {
      setSuccessState(btn);
    })
    .withFailureHandler(function(err, btn) {
      setErrorState(btn, err);
    })
    .saveData(data);
}

function setSavingState(btn) {
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Saving\u2026';
  btn.style.opacity = '0.75';
}

function setSuccessState(btn) {
  btn.innerHTML = '\u2713 Saved!';
  btn.style.opacity = '1';
  btn.style.background = '#137333';
  btn.style.borderColor = '#137333';

  setTimeout(function() {
    btn.disabled = false;
    btn.innerHTML = 'Save';
    btn.style.background = '';
    btn.style.borderColor = '';
    btn.style.opacity = '';
  }, 2500);
}

function setErrorState(btn, err) {
  btn.disabled = false;
  btn.innerHTML = 'Save';
  btn.style.background = '';
  btn.style.borderColor = '';
  btn.style.opacity = '';
  showErrorBanner(err ? err.message : 'Something went wrong. Please try again.');
}
```

**Critical rule:** If the failure handler fires and the button stays disabled, the user has no way to retry. Always restore the button in `setErrorState()` before showing the error.

---

## 9. Status Indicators & Badges

### Inline Badges

Small colored pills used to display status values in text, lists, and table cells.

```css
.badge {
  display: inline-flex;
  align-items: center;
  padding: 2px 8px;
  font-size: 11px;
  font-weight: 500;
  border-radius: 12px;
  white-space: nowrap;
  line-height: 1.6;
}

.badge-success { background: var(--color-success-bg); color: var(--color-success); }
.badge-warning { background: var(--color-warning-bg); color: var(--color-warning); }
.badge-error   { background: var(--color-error-bg);   color: var(--color-error);   }
.badge-info    { background: var(--color-info-bg);    color: var(--color-info);    }
.badge-neutral { background: #f1f3f4;                 color: var(--color-text-secondary); }
```

```html
<span class="badge badge-success">Active</span>
<span class="badge badge-warning">Pending</span>
<span class="badge badge-error">Overdue</span>
<span class="badge badge-neutral">Draft</span>
```

### Status Dots

A 6px colored circle for tight spaces like compact table rows or inline status next to a name.

```css
.status-dot {
  display: inline-block;
  width: 6px;
  height: 6px;
  border-radius: 50%;
  flex-shrink: 0;
  vertical-align: middle;
}

.status-dot.success { background: #34a853; }
.status-dot.warning { background: #fbbc04; }
.status-dot.error   { background: #ea4335; }
.status-dot.neutral { background: #9aa0a6; }
```

```html
<span class="status-dot success"></span> Active
<span class="status-dot warning"></span> Pending
```

---

## 10. Tables in Sidebars

Tables in a 300px sidebar must be compact. Use small font sizes, tight padding, and truncation for long cell values. Never use a table wider than the sidebar — set `width: 100%` and use `table-layout: fixed` when you need strict column widths.

```css
.data-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
  table-layout: auto;
}

.data-table th {
  text-align: left;
  font-weight: 600;
  font-size: 11px;
  color: var(--color-text-secondary);
  padding: 5px 8px;
  border-bottom: 2px solid var(--color-border);
  white-space: nowrap;
}

.data-table td {
  padding: 6px 8px;
  border-bottom: 1px solid var(--color-border);
  color: var(--color-text);
  vertical-align: middle;
}

.data-table tr:last-child td {
  border-bottom: none;
}

.data-table tr:hover td {
  background: var(--color-bg-hover);
}

/* Truncate long text in table cells */
.data-table td.truncate {
  max-width: 100px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
```

```html
<table class="data-table">
  <thead>
    <tr>
      <th>Name</th>
      <th>Status</th>
      <th>Amount</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="truncate">Acme Corp Invoice</td>
      <td><span class="badge badge-success">Paid</span></td>
      <td>$1,250</td>
    </tr>
    <tr>
      <td class="truncate">Riverside Dental</td>
      <td><span class="badge badge-warning">Pending</span></td>
      <td>$840</td>
    </tr>
  </tbody>
</table>
```

---

## 11. Loading States

### Inline Spinner

A pure CSS spinner for use during async operations — inside a button, next to a label, or centered in a loading panel.

```css
.spinner {
  display: inline-block;
  width: 14px;
  height: 14px;
  border: 2px solid var(--color-border);
  border-top-color: var(--color-primary);
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
  flex-shrink: 0;
  vertical-align: middle;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

/* Full-panel loading overlay */
.loading-overlay {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  padding: 32px;
  color: var(--color-text-secondary);
  font-size: 13px;
}
```

```html
<!-- Inline spinner next to text -->
<span class="spinner"></span> Loading…

<!-- Full panel loading state -->
<div class="loading-overlay" id="loadingState">
  <div class="spinner" style="width:28px;height:28px;border-width:3px;"></div>
  <span>Loading data…</span>
</div>
```

### Skeleton Loading

For tables or lists where rows are loading, use skeleton shimmer blocks instead of a full-panel spinner. This reduces perceived wait time by showing the shape of the incoming content.

```css
.skeleton {
  background: linear-gradient(90deg, #f1f3f4 25%, #e8eaed 50%, #f1f3f4 75%);
  background-size: 200% 100%;
  animation: shimmer 1.2s infinite;
  border-radius: var(--border-radius-sm);
  height: 14px;
  display: block;
}

@keyframes shimmer {
  0%   { background-position: 200% 0; }
  100% { background-position: -200% 0; }
}
```

```html
<!-- Skeleton table rows while data loads -->
<tr>
  <td><span class="skeleton" style="width:80%;"></span></td>
  <td><span class="skeleton" style="width:50%;"></span></td>
  <td><span class="skeleton" style="width:60%;"></span></td>
</tr>
```

---

## 12. Error & Success Messages

### Alert Banners

Full-width banners that appear at the top of `.sidebar-body` to communicate operation results. Always inject these above all other content so they appear without scrolling.

```css
.alert {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  padding: 10px 12px;
  border-radius: var(--border-radius-sm);
  font-size: 12px;
  line-height: 1.4;
  margin-bottom: 14px;
}

.alert-success { background: var(--color-success-bg); color: var(--color-success); border: 1px solid #a8d5b5; }
.alert-error   { background: var(--color-error-bg);   color: var(--color-error);   border: 1px solid #f4a7a5; }
.alert-warning { background: var(--color-warning-bg); color: var(--color-warning); border: 1px solid #f5d17a; }
.alert-info    { background: var(--color-info-bg);    color: var(--color-info);    border: 1px solid #a8c4f5; }
```

```javascript
function showErrorBanner(message) {
  const container = document.getElementById('alertContainer');
  container.innerHTML = '<div class="alert alert-error">' + escapeHtml(message) + '</div>';
  container.style.display = 'block';
  document.querySelector('.sidebar-body').scrollTop = 0;
}

function showSuccessBanner(message) {
  const container = document.getElementById('alertContainer');
  container.innerHTML = '<div class="alert alert-success">' + escapeHtml(message) + '</div>';
  container.style.display = 'block';
  setTimeout(function() {
    container.style.display = 'none';
    container.innerHTML = '';
  }, 3000);
}

function clearAlerts() {
  const container = document.getElementById('alertContainer');
  container.style.display = 'none';
  container.innerHTML = '';
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
```

```html
<!-- Place this at the top of .sidebar-body, hidden by default -->
<div id="alertContainer" style="display:none;"></div>
```

### Inline Field Validation Errors

Show field errors directly below the offending input, not in a banner. Add `.input-error` to the input element and show the `.error-text` span.

```css
.input-error {
  border-color: var(--color-error) !important;
  box-shadow: 0 0 0 2px rgba(197, 34, 31, 0.12) !important;
}
```

```html
<div class="field-group">
  <label class="field-label" for="emailField">Email</label>
  <input type="email" id="emailField">
  <span class="error-text" id="emailError">Enter a valid email address.</span>
</div>
```

```javascript
function validateEmail() {
  const input = document.getElementById('emailField');
  const errorEl = document.getElementById('emailError');
  const value = input.value.trim();
  const isValid = value && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);

  input.classList.toggle('input-error', !isValid);
  errorEl.style.display = isValid ? 'none' : 'block';
  return isValid;
}
```

---

## 13. Help Sections & Tooltips

### Collapsible Help Block

Use a native `<details>` element for collapsible help text. No JavaScript required. The `<summary>` is the clickable toggle and the content collapses/expands natively.

```css
.help-block {
  background: var(--color-primary-light);
  border: 1px solid #c5d8f6;
  border-radius: var(--border-radius-sm);
  margin-bottom: 14px;
}

.help-block summary {
  padding: 8px 12px;
  font-size: 12px;
  font-weight: 500;
  color: var(--color-info);
  cursor: pointer;
  list-style: none;
  display: flex;
  align-items: center;
  gap: 6px;
  user-select: none;
}

.help-block summary::-webkit-details-marker { display: none; }

.help-block summary::before {
  content: '?';
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  background: var(--color-info);
  color: #fff;
  font-size: 10px;
  font-weight: 700;
  flex-shrink: 0;
}

.help-block[open] summary::before {
  content: '\00D7';
  font-size: 11px;
}

.help-block .help-content {
  padding: 0 12px 10px;
  font-size: 12px;
  color: var(--color-info);
  line-height: 1.5;
}
```

```html
<details class="help-block">
  <summary>How does this work?</summary>
  <div class="help-content">
    <p>This script reads from column A and writes processed results to column B.
    Make sure column A contains valid email addresses before running.</p>
  </div>
</details>
```

### Inline Help Text

For single-field hints, show help text inline below the label — not in a tooltip. Hover-based tooltips are unreliable inside the sidebar iframe and fail touch accessibility.

```html
<!-- Preferred: always-visible help text below the field -->
<div class="field-group">
  <label class="field-label" for="sheetName">Sheet Name</label>
  <input type="text" id="sheetName" placeholder="Sheet1">
  <span class="help-text">Must match the exact tab name in your spreadsheet, including capitalization.</span>
</div>
```

---

## 14. Scrolling & Sticky Layout

### The Core Rule

The `<body>` must never scroll. The `.sidebar-body` div scrolls instead. This keeps the header and footer permanently visible no matter how much content is in the body.

```css
html, body {
  height: 100%;
  overflow: hidden; /* body never scrolls */
}

.sidebar {
  display: flex;
  flex-direction: column;
  height: 100vh;
}

.sidebar-body {
  flex: 1;
  overflow-y: auto;   /* this is the only scrolling element in the sidebar */
  overflow-x: hidden; /* never allow horizontal scroll */
}
```

### Scroll Behavior & Scrollbar Styling

```css
.sidebar-body {
  scroll-behavior: smooth;
}

/* Slim, unobtrusive scrollbar in WebKit (Chrome) */
.sidebar-body::-webkit-scrollbar {
  width: 6px;
}

.sidebar-body::-webkit-scrollbar-track {
  background: transparent;
}

.sidebar-body::-webkit-scrollbar-thumb {
  background: #dadce0;
  border-radius: 3px;
}

.sidebar-body::-webkit-scrollbar-thumb:hover {
  background: #bdc1c6;
}
```

### Scroll to Top After Errors

When validation fails and errors appear at the top of the form, scroll the body back to the top so the user sees them immediately.

```javascript
function scrollBodyToTop() {
  document.querySelector('.sidebar-body').scrollTop = 0;
}
```

---

## 15. Google Apps Script Integration

### google.script.run — Always Use Both Handlers

Never call `google.script.run` without both `.withSuccessHandler()` and `.withFailureHandler()`. A missing failure handler silently discards errors. The user sees nothing; the button stays disabled forever.

```javascript
// Correct — both handlers always present
google.script.run
  .withSuccessHandler(function(result) {
    // handle success
    console.log('Success:', result);
  })
  .withFailureHandler(function(err) {
    // err.message contains the error string from the server side
    showErrorBanner(err ? err.message : 'An unexpected error occurred.');
  })
  .yourServerFunction(arg1, arg2);
```

### Passing Context to Handlers via withUserObject

Use `.withUserObject()` to pass a DOM reference (like a button) into your success and failure handlers without needing a closure. The object is passed as the second argument to both handlers.

```javascript
google.script.run
  .withUserObject(document.getElementById('saveBtn'))
  .withSuccessHandler(function(result, btn) {
    setSuccessState(btn);
  })
  .withFailureHandler(function(err, btn) {
    setErrorState(btn, err);
  })
  .saveData(data);
```

### Loading Data on Sidebar Open

Pre-populate form fields on open using `window.onload` or `document.addEventListener('DOMContentLoaded', ...)`. A failed data load should be non-blocking — log it but still show the sidebar.

```javascript
window.onload = function() {
  google.script.run
    .withSuccessHandler(function(data) {
      if (!data) return;
      if (data.clientName)  document.getElementById('clientName').value  = data.clientName;
      if (data.clientEmail) document.getElementById('clientEmail').value = data.clientEmail;
      if (data.status)      document.getElementById('statusSelect').value = data.status;
    })
    .withFailureHandler(function(err) {
      // Non-blocking: sidebar still opens, just without pre-filled data
      console.error('Failed to load initial data:', err);
    })
    .getSelectedRowData();
};
```

### Closing the Panel

`google.script.host.close()` works from JavaScript inside the HTML file. It does not work from server-side `.gs` code.

```javascript
// Cancel button
function closePanel() {
  google.script.host.close();
}

// Close from success handler after a save that should dismiss the panel
function setSuccessStateAndClose(btn) {
  btn.innerHTML = '\u2713 Done!';
  setTimeout(function() {
    google.script.host.close();
  }, 1200);
}
```

### Templated HTML — Injecting Server Data on Load

Use `HtmlService.createTemplateFromFile()` when server data is needed immediately on render. This avoids a round-trip `google.script.run` call and prevents the "fields appear empty, then fill in" flash.

```javascript
// Code.gs
function showSidebarWithData() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.clientName = 'Acme Corp';
  template.rowIndex = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  const html = template.evaluate().setTitle('Invoice Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}
```

```html
<!-- Sidebar.html — use <?= ?> for output, <?!= ?> for unescaped HTML -->
<input type="text" id="clientName" value="<?= clientName ?>">
<p>Row: <?= rowIndex ?></p>
```

### Including CSS and JS from Separate Files

Keep styles in `styles.html` and scripts in `scripts.html`. Include them via the `<?!= ?>` scriptlet tag.

```html
<!-- Sidebar.html -->
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <?!= HtmlService.createHtmlOutputFromFile('styles').getContent() ?>
</head>
<body>
  <!-- content -->
  <?!= HtmlService.createHtmlOutputFromFile('scripts').getContent() ?>
</body>
</html>
```

---

## 16. Full Sample Sidebar Template

A complete, copy-paste-ready sidebar with all patterns applied. Inline CSS is used here for a single-file example; in production, split CSS and JS into separate include files.

```html
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
/* ── Reset ─────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body { height: 100%; overflow: hidden; }

/* ── Variables ─────────────────────────────────────── */
:root {
  --color-primary:        #1a73e8;
  --color-primary-dark:   #1557b0;
  --color-primary-light:  #e8f0fe;
  --color-success:        #137333;
  --color-success-bg:     #e6f4ea;
  --color-error:          #c5221f;
  --color-error-bg:       #fce8e6;
  --color-text:           #202124;
  --color-text-secondary: #5f6368;
  --color-text-disabled:  #9aa0a6;
  --color-border:         #dadce0;
  --color-border-focus:   #1a73e8;
  --color-bg-hover:       #f8f9fa;
  --sidebar-padding:      16px;
  --border-radius-sm:     4px;
}

/* ── Base ──────────────────────────────────────────── */
body {
  font-family: 'Google Sans', Roboto, system-ui, sans-serif;
  font-size: 13px;
  line-height: 1.5;
  color: var(--color-text);
  background: #fff;
  -webkit-font-smoothing: antialiased;
}

/* ── Layout ────────────────────────────────────────── */
.sidebar            { display: flex; flex-direction: column; height: 100vh; }

.sidebar-header {
  flex-shrink: 0;
  padding: 16px var(--sidebar-padding) 12px;
  border-bottom: 1px solid var(--color-border);
}

.sidebar-body {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  padding: 16px var(--sidebar-padding);
  scroll-behavior: smooth;
}

.sidebar-footer {
  flex-shrink: 0;
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 12px var(--sidebar-padding);
  border-top: 1px solid var(--color-border);
}

/* Scrollbar */
.sidebar-body::-webkit-scrollbar       { width: 6px; }
.sidebar-body::-webkit-scrollbar-track { background: transparent; }
.sidebar-body::-webkit-scrollbar-thumb { background: #dadce0; border-radius: 3px; }

/* ── Typography ────────────────────────────────────── */
.sidebar-title    { font-size: 16px; font-weight: 500; line-height: 1.3; }
.sidebar-subtitle { font-size: 12px; color: var(--color-text-secondary); margin-top: 2px; }
.section-header   { font-size: 11px; font-weight: 600; letter-spacing: 0.5px; text-transform: uppercase; color: var(--color-text-secondary); margin-bottom: 10px; padding-bottom: 6px; border-bottom: 1px solid var(--color-border); }
.field-label      { display: block; font-size: 12px; font-weight: 500; color: #3c4043; margin-bottom: 4px; }
.help-text        { font-size: 11px; color: var(--color-text-secondary); margin-top: 4px; line-height: 1.4; }
.error-text       { font-size: 11px; color: var(--color-error); margin-top: 4px; display: none; }

/* ── Sections & Fields ─────────────────────────────── */
.section               { margin-bottom: 20px; }
.section:last-child    { margin-bottom: 0; }
.field-group           { margin-bottom: 14px; }
.field-group:last-child { margin-bottom: 0; }

/* ── Form Elements ─────────────────────────────────── */
input[type="text"],
input[type="email"],
select,
textarea {
  width: 100%;
  padding: 7px 10px;
  font-size: 13px;
  font-family: inherit;
  color: var(--color-text);
  background: #fff;
  border: 1px solid var(--color-border);
  border-radius: var(--border-radius-sm);
  outline: none;
  transition: border-color 0.15s, box-shadow 0.15s;
  appearance: none;
  -webkit-appearance: none;
}

input:focus, select:focus, textarea:focus {
  border-color: var(--color-border-focus);
  box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.15);
}

input::placeholder, textarea::placeholder { color: var(--color-text-disabled); }

textarea { resize: vertical; min-height: 72px; }

.select-wrapper          { position: relative; }
.select-wrapper::after   { content:''; position:absolute; right:10px; top:50%; transform:translateY(-50%); border-left:4px solid transparent; border-right:4px solid transparent; border-top:5px solid var(--color-text-secondary); pointer-events:none; }
.select-wrapper select   { padding-right: 28px; cursor: pointer; }

.input-error { border-color: var(--color-error) !important; box-shadow: 0 0 0 2px rgba(197,34,31,0.12) !important; }

/* Checkbox */
.checkbox-item     { display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 8px; }
.checkbox-item:last-child { margin-bottom: 0; }
.checkbox-item input[type="checkbox"] { width:14px; height:14px; flex-shrink:0; accent-color:var(--color-primary); margin:0; cursor:pointer; }
.checkbox-item label { font-size: 13px; cursor: pointer; user-select: none; }

/* ── Buttons ───────────────────────────────────────── */
.btn {
  display: inline-flex; align-items: center; justify-content: center; gap: 6px;
  padding: 7px 16px; font-size: 13px; font-weight: 500; font-family: inherit;
  border-radius: var(--border-radius-sm); border: 1px solid transparent;
  cursor: pointer; white-space: nowrap; min-width: 72px; line-height: 1;
  transition: background 0.15s, box-shadow 0.15s, opacity 0.15s;
  user-select: none;
}
.btn:disabled { opacity: 0.55; cursor: not-allowed; pointer-events: none; }

.btn-primary { background: var(--color-primary); color: #fff; border-color: var(--color-primary); }
.btn-primary:hover:not(:disabled) { background: var(--color-primary-dark); border-color: var(--color-primary-dark); box-shadow: 0 1px 3px rgba(0,0,0,0.2); }

.btn-secondary { background: #fff; color: var(--color-primary); border-color: var(--color-border); }
.btn-secondary:hover:not(:disabled) { background: var(--color-primary-light); border-color: var(--color-primary); }

/* ── Alerts ────────────────────────────────────────── */
.alert         { padding: 10px 12px; border-radius: var(--border-radius-sm); font-size: 12px; line-height: 1.4; margin-bottom: 14px; }
.alert-success { background: var(--color-success-bg); color: var(--color-success); border: 1px solid #a8d5b5; }
.alert-error   { background: var(--color-error-bg);   color: var(--color-error);   border: 1px solid #f4a7a5; }

/* ── Spinner ───────────────────────────────────────── */
.spinner { display:inline-block; width:13px; height:13px; border:2px solid rgba(255,255,255,0.4); border-top-color:#fff; border-radius:50%; animation:spin 0.7s linear infinite; vertical-align:middle; }
@keyframes spin { to { transform: rotate(360deg); } }

/* ── Help block ────────────────────────────────────── */
.help-block         { background: var(--color-primary-light); border: 1px solid #c5d8f6; border-radius: var(--border-radius-sm); margin-bottom: 14px; }
.help-block summary { padding: 8px 12px; font-size: 12px; font-weight: 500; color: #1967d2; cursor: pointer; list-style: none; user-select: none; }
.help-block summary::-webkit-details-marker { display: none; }
.help-block .help-content { padding: 0 12px 10px; font-size: 12px; color: #1967d2; line-height: 1.5; }
</style>
</head>
<body>
<div class="sidebar">

  <!-- ── Header ─────────────────────────────────────── -->
  <div class="sidebar-header">
    <h1 class="sidebar-title">Invoice Generator</h1>
    <p class="sidebar-subtitle">Fill in the details and click Save</p>
  </div>

  <!-- ── Body ───────────────────────────────────────── -->
  <div class="sidebar-body">

    <!-- Alert container (hidden until needed) -->
    <div id="alertContainer" style="display:none;"></div>

    <!-- Section: Client Details -->
    <div class="section">
      <div class="section-header">Client Details</div>

      <div class="field-group">
        <label class="field-label" for="clientName">Client Name</label>
        <input type="text" id="clientName" placeholder="Acme Corp">
        <span class="error-text" id="clientNameError">Client name is required.</span>
      </div>

      <div class="field-group">
        <label class="field-label" for="clientEmail">Email</label>
        <input type="email" id="clientEmail" placeholder="billing@acmecorp.com">
        <span class="help-text">Invoice will be sent to this address.</span>
        <span class="error-text" id="clientEmailError">Enter a valid email address.</span>
      </div>

      <div class="field-group">
        <label class="field-label" for="statusSelect">Status</label>
        <div class="select-wrapper">
          <select id="statusSelect">
            <option value="">— Select —</option>
            <option value="draft">Draft</option>
            <option value="sent">Sent</option>
            <option value="paid">Paid</option>
          </select>
        </div>
      </div>
    </div>

    <!-- Section: Options -->
    <div class="section">
      <div class="section-header">Options</div>

      <div class="field-group">
        <div class="checkbox-item">
          <input type="checkbox" id="optEmail" value="email">
          <label for="optEmail">Send email to client on save</label>
        </div>
        <div class="checkbox-item">
          <input type="checkbox" id="optDrive" value="drive" checked>
          <label for="optDrive">Save PDF to Drive folder</label>
        </div>
      </div>
    </div>

    <!-- Section: Notes -->
    <div class="section">
      <div class="section-header">Notes</div>

      <div class="field-group">
        <label class="field-label" for="notes">Internal Notes</label>
        <textarea id="notes" placeholder="Optional notes for this invoice…"></textarea>
        <span class="help-text">Not visible to the client.</span>
      </div>
    </div>

    <!-- Collapsible help -->
    <details class="help-block">
      <summary>How does this work?</summary>
      <div class="help-content">
        <p>This tool reads the selected row, generates a PDF invoice, and saves it
        to your configured Drive folder. If "Send email" is checked, a copy is
        emailed to the client address entered above.</p>
      </div>
    </details>

  </div><!-- /.sidebar-body -->

  <!-- ── Footer ─────────────────────────────────────── -->
  <div class="sidebar-footer">
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn btn-primary" id="saveBtn" onclick="handleSave()">Save</button>
  </div>

</div><!-- /.sidebar -->

<script>
// ── Save Flow ────────────────────────────────────────
function handleSave() {
  if (!validateForm()) return;

  const btn = document.getElementById('saveBtn');
  clearAlerts();
  setSavingState(btn);

  const data = {
    clientName:  document.getElementById('clientName').value.trim(),
    clientEmail: document.getElementById('clientEmail').value.trim(),
    status:      document.getElementById('statusSelect').value,
    notes:       document.getElementById('notes').value.trim(),
    sendEmail:   document.getElementById('optEmail').checked,
    saveToDrive: document.getElementById('optDrive').checked
  };

  google.script.run
    .withUserObject(btn)
    .withSuccessHandler(function(result, btn) {
      setSuccessState(btn);
    })
    .withFailureHandler(function(err, btn) {
      setErrorState(btn, err);
    })
    .saveInvoiceData(data);
}

// ── Save State Helpers ───────────────────────────────
function setSavingState(btn) {
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Saving\u2026';
  btn.style.opacity = '0.8';
}

function setSuccessState(btn) {
  btn.innerHTML = '\u2713 Saved!';
  btn.style.opacity = '1';
  btn.style.background = '#137333';
  btn.style.borderColor = '#137333';

  setTimeout(function() {
    btn.disabled = false;
    btn.innerHTML = 'Save';
    btn.style.background = '';
    btn.style.borderColor = '';
    btn.style.opacity = '';
  }, 2500);
}

function setErrorState(btn, err) {
  btn.disabled = false;
  btn.innerHTML = 'Save';
  btn.style.background = '';
  btn.style.borderColor = '';
  btn.style.opacity = '';
  showErrorBanner(err ? err.message : 'Something went wrong. Please try again.');
}

// ── Validation ───────────────────────────────────────
function validateForm() {
  let valid = true;

  const clientName  = document.getElementById('clientName');
  const nameError   = document.getElementById('clientNameError');
  if (!clientName.value.trim()) {
    clientName.classList.add('input-error');
    nameError.style.display = 'block';
    valid = false;
  } else {
    clientName.classList.remove('input-error');
    nameError.style.display = 'none';
  }

  const clientEmail = document.getElementById('clientEmail');
  const emailError  = document.getElementById('clientEmailError');
  const emailVal    = clientEmail.value.trim();
  if (emailVal && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailVal)) {
    clientEmail.classList.add('input-error');
    emailError.style.display = 'block';
    valid = false;
  } else {
    clientEmail.classList.remove('input-error');
    emailError.style.display = 'none';
  }

  if (!valid) {
    document.querySelector('.sidebar-body').scrollTop = 0;
  }
  return valid;
}

// ── Alerts ───────────────────────────────────────────
function showErrorBanner(message) {
  const container = document.getElementById('alertContainer');
  container.innerHTML = '<div class="alert alert-error">' + escapeHtml(message) + '</div>';
  container.style.display = 'block';
  document.querySelector('.sidebar-body').scrollTop = 0;
}

function clearAlerts() {
  const container = document.getElementById('alertContainer');
  container.style.display = 'none';
  container.innerHTML = '';
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ── Initial Data Load ────────────────────────────────
window.onload = function() {
  google.script.run
    .withSuccessHandler(function(data) {
      if (!data) return;
      if (data.clientName)  document.getElementById('clientName').value  = data.clientName;
      if (data.clientEmail) document.getElementById('clientEmail').value = data.clientEmail;
      if (data.status)      document.getElementById('statusSelect').value = data.status;
      if (data.notes)       document.getElementById('notes').value = data.notes;
    })
    .withFailureHandler(function(err) {
      // Non-blocking: sidebar still works, just without pre-filled data
      console.error('Failed to load row data:', err);
    })
    .getSelectedRowData();
};
</script>
</body>
</html>
```

---

## 17. What to Avoid

### Layout Mistakes

- **Do not put `overflow-y: auto` on `<body>`** — the header and footer scroll away and the user loses navigation context. Always put scrolling on `.sidebar-body` only. The `<body>` must be `overflow: hidden`.
- **Do not use fixed heights on content containers** — `height: 200px` on a section that can hold 2 or 10 fields will either clip content or leave a white gap. Use `min-height` only when a truly minimum height is needed; otherwise let content flow naturally.
- **Do not use `position: fixed`** — it does not work correctly inside the Google sidebar iframe. Use `flex-shrink: 0` on the header and footer within a flex column layout to achieve sticky behavior.
- **Do not nest scrollable containers** — one `.sidebar-body` with `overflow-y: auto` is the only scrolling element. Nested scroll areas inside a 300px sidebar are confusing and trap users on mobile-style touchpads.
- **Do not set widths wider than the container** — the sidebar is 300px. Setting `width: 320px` on any child element creates horizontal overflow and a broken horizontal scrollbar that looks unprofessional and breaks the layout.

### Form Mistakes

- **Do not omit `box-sizing: border-box`** — an input with `width: 100%` and `padding: 7px 10px` overflows its container if `box-sizing` defaults to `content-box`. This is the single most common cause of broken sidebar layouts.
- **Do not use `<table>` for form layout** — use flexbox or block-flow on `.field-group` wrappers. Tables in a 300px sidebar are nearly impossible to keep from overflowing and require constant column-width maintenance.
- **Do not use `placeholder` as a label substitute** — placeholders disappear on input and are inaccessible to screen readers. Always include a visible `<label>` element associated with its input via `for` / `id`.
- **Do not skip `font-family: inherit` on inputs** — Chrome's default input font inside the sidebar iframe differs from the body font. Without explicit inheritance, inputs render in a different typeface than surrounding labels and text.
- **Do not use `<select>` without a wrapper div** — without `.select-wrapper`, you cannot add the custom chevron after removing the native arrow with `appearance: none`. Bare `<select>` elements with `appearance: none` look arrow-less and broken.

### Google Apps Script Mistakes

- **Do not call `google.script.run` without `.withFailureHandler()`** — errors are silently discarded. The user sees nothing; the button stays in the "Saving…" state forever. Both handlers are mandatory.
- **Do not pass DOM elements to server functions** — `google.script.run` serializes arguments to JSON before sending them to the server. DOM elements, functions, and `undefined` values are stripped. Pass only plain objects, strings, arrays, and numbers.
- **Do not call `google.script.host.close()` from server-side `.gs` code** — it only works from client-side JavaScript inside the HTML file. Call it from a `withSuccessHandler` or a button's `onclick`.
- **Do not use `document.write()`** — it wipes the entire sidebar document, destroying all CSS and layout. Use `innerHTML` on a target element or `createElement`/`appendChild` for DOM injection.
- **Do not make repeated `google.script.run` calls for data that does not change** — Apps Script server calls have 300ms–2s latency. Cache data in JavaScript variables after the first load. Re-fetch only when the underlying data may have changed.
- **Do not open a sidebar without a title** — `setTitle()` on the `HtmlOutput` object sets the sidebar header label shown by Google. An untitled sidebar appears blank in that space, which looks broken.

### Visual Mistakes

- **Do not use more than 3 font sizes in a single sidebar** — label (12px), body (13px), and help text (11px) is the complete type scale. Adding additional sizes breaks the visual hierarchy and makes the sidebar feel unpolished.
- **Do not hard-code color hex values in element styles** — always use CSS variables. Hard-coded colors scatter through the stylesheet and make theme adjustments require a codebase-wide search-and-replace.
- **Do not use `<hr>` for section dividers** — style `border-bottom` on the `.section-header` instead. Bare `<hr>` elements use inconsistent browser-default styling and look cheap at small sizes.
- **Do not animate `height`, `width`, or `padding`** — animating layout-affecting properties causes jank inside the iframe. Animate only `opacity`, `transform`, and `background-color`, which use the compositor and do not cause reflow.
- **Do not leave buttons disabled after a failure** — if a save fails and the button stays in the "Saving…" state, the user has no way to retry and must close and reopen the sidebar. Always call `setErrorState()` in the failure handler to restore the button.

---

*This document is part of the TAKScripts Polish Series. See also: POLISH_DRIVE.md, POLISH_SHEETS.md*
