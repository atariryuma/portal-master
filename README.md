# Portal Master (Google Apps Script)

Portal Master is a Google Sheets + Apps Script toolkit for school operations.
It provides menu-driven workflows for annual event import, daily operations,
duty assignment, module learning management, hour aggregation, calendar sync,
and annual rollover tasks.

## What This Project Includes

- Menu-driven operations from `onOpen()` (`menu.js`)
- Annual event import/update workflows
- Daily utilities (jump to today, PDF export, open report folder)
- Duty assignment and duty-only updates
- Cumulative hour calculation and grade-based hour aggregation
- Module planning/actual tracking integration
- Trigger setting UI and scheduling logic

## Environment

- Runtime: Google Apps Script (V8)
- Time zone: `Asia/Tokyo` (`appsscript.json`)
- Advanced service: Drive API v3 enabled (`appsscript.json`)
- Local sync tool: `clasp`

## Required Sheets

At minimum, these sheets must exist:

- `マスター`
- `app_config`
- `時数様式`
- `年間行事予定表` (required for schedule-based operations)

## File Structure

### Core

| File | Description |
| ---- | ----------- |
| `menu.js` | Menu entry points (`onOpen`) |
| `common.js` | Shared constants, utilities, config access |
| `testRunner.js` | Quick/Full test suite (60+ tests) |

### Operations

| File | Description |
| ---- | ----------- |
| `assignDuty.js` | Duty assignment (batch-optimized) |
| `updateAnnualEvents.js` | Annual event update from master sheet |
| `updateAnnualDuty.js` | Duty-only column update |
| `importAnnualEvents.js` | Annual events import from external spreadsheet |
| `aggregateSchoolEvents.js` | Grade-based hour aggregation |
| `calculateCumulativeHours.js` | Cumulative hour calculation |
| `syncCalendars.js` | Google Calendar sync |
| `saveToPDF.js` | Weekly report PDF export |
| `countDutyStars.js` | Vacation duty star counting |
| `setDailyHyperlink.js` | Today's date navigation link |
| `openWeeklyReportFolder.js` | Weekly report folder opener |

### Module Learning (split from single 2,400-line file)

| File | Description |
| ---- | ----------- |
| `moduleHoursConstants.js` | Module learning constants |
| `moduleHoursDialog.js` | Dialog/UI interaction |
| `moduleHoursPlanning.js` | Cycle plan allocation algorithms |
| `moduleHoursControl.js` | Sheet I/O, migration, settings |
| `moduleHoursDisplay.js` | Formatting, utilities, cumulative output |

### Configuration & Triggers

| File                       | Description                              |
| -------------------------- | ---------------------------------------- |
| `triggerSettings.js`       | Trigger settings and rebuild logic       |
| `annualUpdateSettings.js`  | Annual update settings validation/save   |

### UI Dialogs (HTML)

| File | Description |
| ---- | ----------- |
| `DateSelector.html` | Date selection dialog |
| `modulePlanningDialog.html` | Module learning management dialog |
| `triggerSettingsDialog.html` | Trigger settings dialog |

## Constants Design

All magic numbers are centralized in `common.js`:

| Constant Group | Purpose |
| -------------- | ------- |
| `MASTER_SHEET` | Master sheet structure (columns, rows) |
| `DUTY_ROSTER_SHEET` | Duty roster sheet columns |
| `ANNUAL_SCHEDULE` | Annual schedule sheet columns |
| `JISUU_TEMPLATE` | Hour template layout |
| `WEEKLY_REPORT` | Weekly report PDF settings |
| `CUMULATIVE_SHEET` | Cumulative hours sheet |
| `IMPORT_CONSTANTS` | Import configuration |
| `CONFIG_CELLS` | Settings sheet cell addresses |

## Local Workflow (clasp)

1. Authenticate

```bash
clasp login
```

1. Pull latest script (if needed)

```bash
clasp pull
```

1. Push local changes

```bash
clasp push
```

1. Open Apps Script editor (optional)

```bash
clasp open
```

## Testing

Testing is executed from `testRunner.js`. The suite contains 63 tests across 7 groups.

### Quick Test (daily development)

- Run: `runQuickTest()`
- Purpose: catch breakages quickly during iterative edits

### Full Test (before release)

- Run: `runAllTests()`
- Purpose: broad regression check before production use
- Required when shipping changes or modifying core logic

### Test Groups

1. Environment checks (spreadsheet, sheets)
1. Settings and configuration
1. Data integrity (master sheet, calendar)
1. Public function definitions
1. Trigger configuration
1. Module learning integration
1. Code quality (constants, var check, log prefixes, error handling, XSS, batch reads, decomposition)

### Failure Handling

- If `runAllTests()` has any failure, do not release.
- Fix code (or test if the test is incorrect), then rerun.
- Use the `【エラー詳細】` section in logs as the source of truth.

### Test Design Rules

- Prefer behavior checks (input/output/side effects), not symbol existence checks.
- For new features, add:
  - at least 1 normal-path test
  - at least 1 error-path test

## Recommended Change Flow

1. Implement change
1. Run `runQuickTest()`
1. Run `runAllTests()` before release
1. `clasp push`
1. Re-run critical operation once in the target spreadsheet

## Notes

- This repository contains Apps Script source; execution still depends on
  spreadsheet data quality and Google service permissions.
- Some workflows (Drive/Calendar/PDF) depend on external service availability
  and account permissions.
- All files share GAS global scope; no module imports needed between files.
