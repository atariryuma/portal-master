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

## Local Workflow (clasp)
1. Authenticate
```bash
clasp login
```
2. Pull latest script (if needed)
```bash
clasp pull
```
3. Push local changes
```bash
clasp push
```
4. Open Apps Script editor (optional)
```bash
clasp open
```

## Testing Policy
Testing is executed from `testRunner.js`.

### 1) Daily Development (fast check)
- Run: `runQuickTest()`
- Purpose: catch breakages quickly during iterative edits
- Use when:
  - making small changes
  - checking local stability before frequent pushes

### 2) Before Release (full validation)
- Run: `runAllTests()`
- Purpose: broad regression check before production use
- Required when:
  - shipping changes to users
  - changing core logic (aggregation, settings, module, triggers, UI flow)

### 3) Failure Handling
- If `runAllTests()` has any failure, do not release.
- Fix code (or test if the test is incorrect), then rerun.
- Use the `【エラー詳細】` section in logs as the source of truth.

### 4) Test Design Rules
- Prefer behavior checks (input/output/side effects), not symbol existence checks.
- For new features, add:
  - at least 1 normal-path test
  - at least 1 error-path test

## Recommended Change Flow
1. Implement change
2. Run `runQuickTest()`
3. Run `runAllTests()` before release
4. `clasp push`
5. Re-run critical operation once in the target spreadsheet

## Main Files
- `menu.js`: menu entry points
- `common.js`: shared constants/utilities
- `aggregateSchoolEvents.js`: grade-based hour aggregation
- `moduleHours.js`: module planning/actual integration
- `triggerSettings.js`: trigger settings and trigger rebuild logic
- `DateSelector.html`, `modulePlanningDialog.html`, `triggerSettingsDialog.html`: UI dialogs
- `testRunner.js`: quick/full test suite

## Notes
- This repository contains Apps Script source; execution still depends on
  spreadsheet data quality and Google service permissions.
- Some workflows (Drive/Calendar/PDF) depend on external service availability
  and account permissions.
