# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Portal Master is a Google Apps Script (GAS) project for school operations management. It runs inside Google Sheets with the V8 runtime and uses `clasp` for local development sync.

## Commands

```bash
# Push local changes to Apps Script
clasp push

# Pull remote changes
clasp pull

# Open Apps Script editor in browser
clasp open
```

### Testing (executed in Apps Script editor, not locally)

- **Quick test** (daily development): Run `runQuickTest()` in the Apps Script editor
- **Full test** (before release): Run `runAllTests()` in the Apps Script editor — 60+ tests across 7 phases
- Tests cannot be run locally. They execute inside the GAS environment against live spreadsheet data.
- If any test fails, check `【エラー詳細】` in the Apps Script execution log.

## Architecture

### GAS Global Scope

All `.js` files share a single global scope — there are no module imports. File load order is non-deterministic, so:
- Constants referenced across files must be defined in the same file or in `common.js`
- `CUMULATIVE_EVENT_CATEGORIES` is defined in `common.js` (not `calculateCumulativeHours.js`) because it depends on `EVENT_CATEGORIES`
- `normalizeToDate()` is defined in `moduleHoursDisplay.js` — functions in `common.js` cannot reference it at top level due to load-order constraints

### Centralized Constants (`common.js`)

All spreadsheet structure (column indices, cell addresses, sheet names) is defined as `Object.freeze()`-protected constants in `common.js`. Key groups:
- `MASTER_SHEET`, `ANNUAL_SCHEDULE`, `JISUU_TEMPLATE` — sheet column/row mappings
- `CONFIG_CELLS`, `TRIGGER_CONFIG_CELLS`, `ANNUAL_UPDATE_CONFIG_CELLS` — `app_config` sheet cell addresses
- `SCHEDULE_COLUMNS` — 0-based column indices for the annual schedule data range
- `EVENT_CATEGORIES` — category name → abbreviation mapping used in hour counting

When adding new constants, always wrap with `Object.freeze()` to prevent accidental mutation.

### Module Learning Subsystem (5 files)

Originally a single 2,400-line file, decomposed into:
- `moduleHoursConstants.js` — constants (cycles, headers, markers, layout defaults)
- `moduleHoursControl.js` — sheet I/O, layout management, legacy migration
- `moduleHoursPlanning.js` — cycle plan allocation, daily plan generation, school day mapping
- `moduleHoursDialog.js` — dialog UI state and server-side handlers
- `moduleHoursDisplay.js` — formatting, cumulative output, display column management, date/number utilities (`normalizeToDate`, `isNonEmptyCell`, `toNumberOrZero`, `formatMonthKey`, `formatInputDate`)

All module data lives in a single `module_control` sheet (migrated from multi-sheet structure). Settings use `PropertiesService.getDocumentProperties()` with `MODULE_` prefix.

**Weekday filter:** `extractSchoolDayRows()` filters school days by user-configured weekdays (`MODULE_WEEKDAYS_ENABLED` in PropertiesService, default `[1,3,5]` = Mon/Wed/Fri). `MODULE_WEEKDAY_PRIORITY` controls allocation order within a week (Mon→Wed→Fri→Tue→Thu). `getEnabledWeekdays()` reads the setting; `serializeWeekdays()` writes it.

### Data Flow Pattern

`app_config` sheet (cell addresses in `CONFIG_CELLS` / `TRIGGER_CONFIG_CELLS`) → read by `getSettingsSheetOrThrow()` → used by all features that need folder IDs, calendar IDs, or trigger configuration.

### Dialog Pattern

HTML dialogs (`*.html`) call server-side functions via `google.script.run`. Each dialog has a paired `.js` file providing the server-side API (e.g., `triggerSettingsDialog.html` ↔ `triggerSettings.js`).

## Coding Conventions

### Syntax Rules

- **`function` keyword only** — no arrow functions, no shorthand method syntax (consistency across all GAS files)
- **`const`/`let` only** — no `var`
- **Strict equality** — always `===`/`!==`, use `Number()` for explicit coercion when comparing sheet values
- **No `eval()`** — use explicit object maps for dynamic references
- **No `for...in` / `for...of`** — use `Object.keys().forEach()` or `Array.prototype.forEach()` for iteration
- **Explicit object literals** — always use `{ key: value }` form, never ES6 shorthand `{ key }` (readability and consistency)
- **`parseInt` with radix** — always `parseInt(value, 10)` to prevent unexpected octal parsing

### GAS API Performance

- **Batch GAS API calls** — never loop individual `getValue()`/`setValue()`; always use `getRange().getValues()` / `setValues()` for ranges
- **Safe file operations** — when replacing files (e.g., PDF export), create the new file first, then delete the old file (prevents data loss on creation failure)
- **No DEBUG logs in production** — `Logger.log('[DEBUG] ...')` is for development only; remove before committing

### Date Handling

- **Use `normalizeToDate()`** (defined in `moduleHoursDisplay.js`) instead of `new Date(value)` for parsing date values from sheets or user input
- **Exception:** `common.js` functions cannot depend on `normalizeToDate()` due to GAS load-order constraints — use explicit date parsing inline instead
- **Use `formatMonthKey()`** instead of raw `Utilities.formatDate()` for "yyyy-MM" keys
- **Use `formatInputDate()`** for "yyyy-MM-dd" keys

### Error Handling & UI

- **Error display** — `showAlert(message, title)` for user-facing errors; `Logger.log('[LEVEL] ...')` for server-side logging with `[INFO]`, `[WARNING]`, `[ERROR]` prefixes
- **Confirmation dialogs** — use `ui.ButtonSet.OK_CANCEL` (not `OK`) for destructive or irreversible actions, and check the user's response
- **Japanese UI text** — all user-facing strings are in Japanese

### Aggregation Pattern

- When counting categories by abbreviation, build a reverse lookup map (`abbreviation → categoryName`) before the data loop to achieve O(n) instead of O(n*m) nested loops. See `calculateCumulativeHours.js:calculateResultsForGrade` for the reference implementation.

## Required Sheets

Operations depend on these sheet names: `マスター`, `app_config`, `時数様式`, `年間行事予定表`, `累計時数`, `日直表`, `module_control`.

## Domain Knowledge (Business Rules)

### Fiscal Year (年度)

Japanese school fiscal year runs April 1 – March 31. Dates in January–March belong to the **previous** fiscal year (e.g., 2026-01-15 → FY2025). `getFiscalYear()` in `moduleHoursDisplay.js` implements this.

### Cumulative Hours (累計時数)

- `○` in a schedule cell = 1 regular class hour
- Category abbreviations (e.g., `儀式`, `文化`) = 1 special activity hour, counted per category
- `補習` (supplementary lessons) is **excluded** from cumulative totals because it falls outside the standard class hours defined by MEXT (文部科学省)

### Module Learning (モジュール学習)

- 1 session = 15 minutes (smallest unit of module learning)
- 1 コマ = 45 minutes = 3 sessions (standard class hour equivalent)
- Sessions are allocated to school days using **weekday priority**: Mon → Wed → Fri → Tue → Thu. This creates an alternating-day pattern for optimal short-duration learning distribution.
- Default plan: 4 cycles/year × 7 コマ/cycle = 28 コマ (21 hours)

### Annual Update (年度更新)

The `copyAndClear` function copies the current file as a backup, then clears the **current** file (not the copy). This preserves the original file's URL so bookmarks and shared links remain valid.

### Dialog Pattern

HTML dialogs use `createTemplateFromFile().evaluate()` to support shared CSS includes via `<?!= include_('dialogStyles') ?>`. The `include_()` helper is defined in `common.js`. All 4 dialogs (triggerSettings, annualUpdateSettings, DateSelector, modulePlanning) use this pattern.

### Module Learning Settings

Weekday priority for module learning allocation is stored in `PropertiesService` as `MODULE_WEEKDAY_PRIORITY` (JSON array of day numbers, e.g. `[1,3,5,2,4]`). The `weekdayPriority()` function in `moduleHoursPlanning.js` loads from settings with a per-execution cache (`activeWeekdayPriorityMap_`). Call `resetWeekdayPriorityCache_()` after saving new priority values.

## Testing Conventions

- Behavior checks (input/output/side effects) over symbol existence checks
- At least 1 normal-path + 1 error-path test per feature
- Tests must be **non-destructive** — use temporary sheets with `finally` cleanup, or read-only verification of existing data. Never write to production sheets in tests.
- Test groups: Environment → Module Integration → Data Processing → Settings → Common Functions → Operational Workflows → Code Quality
