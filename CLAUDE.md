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

- **Quick test** (daily development): Run `runQuickTest()` — 9 critical tests across 2 groups
- **Full test** (before release): Run `runAllTests()` — 73 tests across 8 phases
- Tests cannot be run locally. They execute inside the GAS environment against live spreadsheet data.
- Tests can also be run remotely via the web app endpoint: `?page=tests&suite=full&phase=N` (N=1-8)
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
- `DUTY_ROSTER_SHEET` — duty roster sheet column mappings
- `WEEKLY_REPORT` — weekly report sheet names, cell addresses, row ranges, `PDF_OPTIONS`
- `CUMULATIVE_SHEET` — cumulative hours sheet name, grade start row, date cell
- `IMPORT_CONSTANTS` — import row count and source sheet name
- `CONFIG_CELLS`, `TRIGGER_CONFIG_CELLS`, `ANNUAL_UPDATE_CONFIG_CELLS` — `app_config` sheet cell addresses
- `SCHEDULE_COLUMNS` — 0-based column indices for the annual schedule data range
- `WEEKDAY_MAP` — day number to `ScriptApp.WeekDay` mapping
- `EVENT_CATEGORIES` — category name → abbreviation mapping used in hour counting
- `CUMULATIVE_EVENT_CATEGORIES` — derived from `EVENT_CATEGORIES` (excludes `補習`)
- `MODULE_SHEET_NAMES` — module-related sheet names (control, plan summary)
- `MODULE_SETTING_KEYS` — PropertiesService key names with `MODULE_` prefix
- `MODULE_CUMULATIVE_COLUMNS` — column indices for module output in cumulative sheet
- `CALENDAR_MANAGED_DESCRIPTION_MARKER` — managed calendar identification string

When adding new constants, always wrap with `Object.freeze()` to prevent accidental mutation.

### File Organization

**JS files (24):**

| File | Purpose |
|------|---------|
| `common.js` | Centralized constants, date/name/sheet utilities, HTML template helper (`include_()`) |
| `menu.js` | Menu structure (`onOpen`), internal sheet visibility management |
| `importAnnualEvents.js` | Imports annual event data from external spreadsheet to master sheet |
| `updateAnnualEvents.js` | Reflects master sheet data to annual schedule sheet (batch write) |
| `updateAnnualDuty.js` | Syncs only duty column from master to annual schedule |
| `assignDuty.js` | Auto-assigns duty roster to master sheet using round-robin |
| `countDutyStars.js` | Counts star marks (☆) in annual schedule for duty staff |
| `calculateCumulativeHours.js` | Cumulative hours per grade up to nearest Saturday, triggers module sync |
| `aggregateSchoolEvents.js` | Grade-specific class hour aggregation (low/mid/high groups), date selector handler |
| `syncCalendars.js` | Google Calendar sync with diff-based update, holiday skip, managed event markers |
| `saveToPDF.js` | Weekly report PDF export to Google Drive, safe file replacement |
| `copyAndClear.js` | Annual update: backup copy + current file data clear (preserves URL) |
| `triggerSettings.js` | Trigger settings dialog server-side API, trigger creation/deletion |
| `annualUpdateSettings.js` | Annual update settings dialog server-side API |
| `setDailyHyperlink.js` | Sets today's date hyperlink in B1 of annual schedule |
| `openWeeklyReportFolder.js` | Opens weekly report Drive folder via modal dialog |
| `moduleHoursConstants.js` | Module learning constants (headers, markers, layout defaults, weekday labels/priority) |
| `moduleHoursControl.js` | Module control sheet I/O, layout management, per-execution caches, settings via PropertiesService |
| `moduleHoursPlanning.js` | Session allocation algorithms (annual/monthly modes), school day mapping, exception handling |
| `moduleHoursDialog.js` | Module planning dialog state assembly, server-side handlers for dialog actions |
| `moduleHoursDisplay.js` | Cumulative output, formatting, plan summary sheet, date/number utilities |
| `webapp.js` | Web app entry points (doGet/doPost), diagnostics, test runner endpoint, orphan trigger cleanup |
| `testRunner.js` | Test suite: 73 tests (full) / 9 tests (quick) |

**HTML files (7):**

| File | Purpose |
|------|---------|
| `triggerSettingsDialog.html` | Trigger configuration dialog UI |
| `annualUpdateSettingsDialog.html` | Annual update settings dialog UI |
| `modulePlanningDialog.html` | Module learning management dialog UI |
| `DateSelector.html` | Date range picker for grade aggregation |
| `dialogStyles.html` | Shared CSS styles (included via `include_()`) |
| `userGuide.html` | Comprehensive user guide |
| `webappHome.html` | Web app dashboard UI (endpoint list, status/diagnostics/test buttons) |

### Module Learning Subsystem (5 files)

Originally a single 2,400-line file, decomposed into:
- `moduleHoursConstants.js` — constants (headers, markers, layout defaults, weekday labels/priority)
- `moduleHoursControl.js` — sheet I/O, layout management, per-execution caches, settings via PropertiesService
- `moduleHoursPlanning.js` — session allocation (annual/monthly modes), daily plan generation, school day mapping, exception handling
- `moduleHoursDialog.js` — dialog UI state and server-side handlers
- `moduleHoursDisplay.js` — formatting, cumulative output, display column management, date/number utilities (`normalizeToDate`, `isNonEmptyCell`, `toNumberOrZero`, `formatMonthKey`, `formatInputDate`, `getFiscalYear`)

All module data lives in a single `module_control` sheet (migrated from multi-sheet structure). Settings use `PropertiesService.getDocumentProperties()` with `MODULE_` prefix.

**Weekday filter:** `extractSchoolDayRows()` filters school days by user-configured weekdays (`MODULE_WEEKDAYS_ENABLED` in PropertiesService, default `[1,3,5]` = Mon/Wed/Fri). `MODULE_WEEKDAY_PRIORITY` in `moduleHoursConstants.js` defines allocation order (Mon→Wed→Fri→Tue→Thu). `getEnabledWeekdays()` reads the setting; `serializeWeekdays()` writes it. `weekdayPriority()` in `moduleHoursPlanning.js` resolves priority from the frozen constant.

**Settings storage:** `PropertiesService.getDocumentProperties()` with `MODULE_` prefix. Key names defined in `MODULE_SETTING_KEYS` (`common.js`). Read by `readModuleSettingsMap()`, written by `upsertModuleSettingsValues()`. Plan summary sheet: `MODULE_SHEET_NAMES.PLAN_SUMMARY` = `'モジュール学習計画'`, generated by `writeModulePlanSummarySheet()` in `moduleHoursDisplay.js`.

### Data Flow Pattern

Two-tier settings architecture:
1. **Sheet-based** (`app_config`): Cell addresses in `CONFIG_CELLS` / `TRIGGER_CONFIG_CELLS` / `ANNUAL_UPDATE_CONFIG_CELLS` → read by `getSettingsSheetOrThrow()` → folder IDs, calendar IDs, trigger configuration
2. **PropertiesService-based** (`MODULE_*` prefix): Module-specific settings (weekdays, planning range, timestamps) → read by `readModuleSettingsMap()` → module learning subsystem

### Dialog Pattern

HTML dialogs (`*.html`) call server-side functions via `google.script.run`. Each dialog has a paired `.js` file providing the server-side API:
1. `triggerSettingsDialog.html` ↔ `triggerSettings.js`
2. `annualUpdateSettingsDialog.html` ↔ `annualUpdateSettings.js`
3. `modulePlanningDialog.html` ↔ `moduleHoursDialog.js`
4. `DateSelector.html` ↔ `aggregateSchoolEvents.js`

For shared styles, render dialogs with `createTemplateFromFile(...).evaluate()` and include `<?!= include_('dialogStyles') ?>` (helper: `include_()` in `common.js`). This applies to trigger settings, annual update settings, and DateSelector dialogs. `modulePlanningDialog.html` does **not** use shared styles (it uses inline CSS).

### Web App

`webapp.js` provides `doGet(e)` and `doPost(e)` entry points for the web app deployment. Available endpoints:
- `?page=home` — dashboard UI (`webappHome.html`)
- `?page=status` — detailed system status (JSON: spreadsheet info, sheets, triggers, module settings, app_config)
- `?page=diagnostics` — health checks (required sheets, key functions, trigger list)
- `?page=tests&suite=full&phase=N` — run test suite remotely (phase 1-8, or omit phase for all)
- POST `{"action": "ping"}` — connectivity check
- POST `{"action": "deleteOrphanTriggers"}` — remove triggers with no matching function

`isTriggeredExecution_()` in `common.js` detects whether code is running from a time trigger or web app (no UI context) vs. menu/dialog (UI available). Used by `syncCalendars()` to skip confirmation dialogs during trigger execution.

### Per-Execution Caching

`moduleHoursControl.js` uses three per-execution caches (`let cache_ = null` at file top level). GAS discards these automatically when execution ends.

- `moduleHoursSheetsCache_` — caches initialized `module_control` sheet object
- `moduleSettingsMapCache_` — caches PropertiesService read of `MODULE_*` keys; returns shallow copy to prevent mutation; invalidated by `upsertModuleSettingsValues()`
- `moduleControlLayoutCache_` — caches section marker row positions (plan/exceptions); invalidated by `invalidateModuleControlLayoutCache_()` after row insertions

Pattern: always invalidate after write operations that change the cached state.

### Concurrency Control

`LockService.getDocumentLock()` with `tryLock(10000)` is used in:
- `syncCalendars()` — prevents concurrent calendar sync
- `copyAndClear()` — prevents concurrent annual update

Both release the lock in a `finally` block. Show user-facing error on lock failure.

### Menu Structure

`onOpen()` in `menu.js` builds the following menu hierarchy:

```
🎯 ポータルマスター
├── 🚀 導入: importAnnualEvents, updateAnnualEvents
├── ⚙️ 設定: showAnnualUpdateSettingsDialog, showTriggerSettingsDialog
├── 📅 日常業務: setDailyHyperlink, saveToPDF, openWeeklyReportFolder
├── 👥 日直: assignDuty, updateAnnualDuty, countStars
├── 📊 集計: calculateCumulativeHours, aggregateSchoolEventsByGrade, showModulePlanningDialog
├── 🔁 連携と年度更新: syncCalendars, copyAndClear
└── ❓ ヘルプ: showUserGuide, showCreatorInfo
```

`onOpen` also calls `hideInternalSheetsForNormalUse_()` which hides `module_control` and `app_config`. Master sheet is NOT hidden on initial load (users edit it during setup); it is hidden after `updateAnnualEvents()` completes.

## Coding Conventions

### Syntax Rules

- **`function` keyword only** — no arrow functions, no shorthand method syntax (consistency across all GAS files)
- **`const`/`let` only** — no `var`
- **Default parameters OK** — e.g., `function showAlert(message, title = '通知')` (GAS V8 runtime supports this)
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

Core operations depend on these sheet names:

| Sheet | Purpose |
|-------|---------|
| `マスター` | Source data: events, attendance, duty assignments |
| `app_config` | Settings storage (folder IDs, calendar IDs, trigger config) |
| `時数様式` | Template sheet for grade aggregation output |
| `年間行事予定表` | Annual schedule: dates, events, attendance, duty |
| `累計時数` | Cumulative hours output per grade |
| `日直表` | Duty roster: names and duty numbers |
| `module_control` | Module learning data (V4 format, hidden from users) |
| `週報（時数あり）` | Weekly report (current week) |
| `週報（時数あり）次週` | Weekly report (next week) |

Auto-generated sheets:
- `モジュール学習計画` — module plan summary (created by `writeModulePlanSummarySheet()`)
- Grade aggregation output sheets — created by `aggregateSchoolEventsByGrade()` using `時数様式` as template

## Domain Knowledge (Business Rules)

### Fiscal Year (年度)

Japanese school fiscal year runs April 1 – March 31. Dates in January–March belong to the **previous** fiscal year (e.g., 2026-01-15 → FY2025). `getFiscalYear()` in `moduleHoursDisplay.js` implements this. `MODULE_FISCAL_YEAR_START_MONTH = 4` in `common.js`.

### Cumulative Hours (累計時数)

- `○` in a schedule cell = 1 regular class hour
- Category abbreviations (e.g., `儀式`, `文化`) = 1 special activity hour, counted per category
- `補習` (supplementary lessons) is **excluded** from cumulative totals because it falls outside the standard class hours defined by MEXT (文部科学省)
- Calculation endpoint: nearest Saturday from today (`getCurrentOrNextSaturday()`)
- After cumulative hours calculation, `syncModuleHoursWithCumulative()` writes module learning data to the cumulative sheet

### Module Learning (モジュール学習)

- 1 session = 15 minutes (smallest unit of module learning)
- 1 コマ = 45 minutes = 3 sessions (standard class hour equivalent)
- Sessions are allocated to school days using configurable **enabled weekdays** (default: Mon/Wed/Fri) with weekday priority distribution
- **V4 data format**: 1 row per grade per fiscal year (6 grades × 1 row = 6 rows/year), with 12 monthly columns (m4–m3, April–March). Each grade independently chooses `annual` mode (single total) or `monthly` mode (per-month targets)
- Default plan: 28 コマ/year per grade (`MODULE_DEFAULT_ANNUAL_KOMA`)
- **Reserve/deficit tracking**: `reserveByGrade` = available school days − planned sessions. Displayed in cumulative sheet and dialog
- **Exception handling**: date-based session adjustments stored in EXCEPTIONS_TABLE section of `module_control`. Applied by `applyModuleExceptions()` in `moduleHoursPlanning.js`
- **Migration chain**: V1 (multi-sheet) → V2 (single control, cycle-based) → V3 (annual 8-column) → V4 (per-grade 17-column) with auto-detection

### Grade Aggregation (学年別授業時数集計)

- Entry point: `aggregateSchoolEventsByGrade()` shows `DateSelector.html` for date range input
- Processes grades in groups: low (1,2), mid (3,4), high (5,6)
- Creates output sheets by copying layout from `時数様式` template
- Integrates module learning data (MOD column) with fallback to preserved existing values via `captureExistingModValuesByMonth()` on calculation failure
- Uses the same reverse-lookup aggregation pattern as `calculateCumulativeHours`

### Calendar Sync (カレンダー同期)

- Two separate calendars: event calendar (校内行事 from column D) and external calendar (対外行事 from column M)
- **Managed markers**: `CALENDAR_MANAGED_DESCRIPTION_MARKER` (`[PORTAL_MASTER_CALENDAR]`) in calendar description, `CALENDAR_SYNC_MANAGED_MARKER` (`[PORTAL_MASTER_MANAGED]`) in event description — only managed events are deleted during sync
- **Diff-based sync**: builds event key from `[title, startTime, endTime]`, creates/deletes only changed events
- **Holiday skip**: detects Japanese holidays calendar and skips syncing on holidays
- **Rate limiting**: `Utilities.sleep(200)` after event change batches
- **Time parsing**: supports range times (e.g., "9:00~12:00"), single times, all-day events; handles full-width characters via `convertFullWidthToHalfWidth()`

### PDF Export (週報PDF保存)

- Exports both weekly report sheets to PDF in Google Drive
- **Row height adjustment**: toggles row heights based on trigger cell (`U41`) content presence to show relevant week section
- **Safe file replacement**: creates new file first, collects old files with matching name, then trashes old files
- **File naming**: `{names from B1:D1}({start date}~{end date}).pdf`
- Folder auto-created if missing; folder ID cached in `app_config` C14

### Annual Update (年度更新)

The `copyAndClear` function copies the current file as a backup, then clears the **current** file (not the copy). This preserves the original file's URL so bookmarks and shared links remain valid. Backup integrity is verified (sheet existence + data check) before clearing. Uses `LockService` to prevent concurrent execution.

### Duty System (日直)

- **Assignment** (`assignDuty`): reads duty roster (names + duty numbers), groups by number, extracts first name via `extractFirstName()`, writes to master sheet AO column in round-robin order
- **Annual sync** (`updateAnnualDuty`): syncs master AO column to annual schedule R column
- **Star counting** (`countStars`): reads annual schedule R column (format: "☆☆\n太郎\n花子"), counts ☆ per duty person, writes totals to duty roster E column

## Testing Conventions

- Behavior checks (input/output/side effects) over symbol existence checks
- At least 1 normal-path + 1 error-path test per feature
- Tests must be **non-destructive** — use temporary sheets with `finally` cleanup, or read-only verification of existing data. Never write to production sheets in tests.
- Test groups: Environment → Module Integration → Data Processing → Settings → Common Functions → Operational Workflows → Code Quality
