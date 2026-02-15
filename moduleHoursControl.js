/**
 * @fileoverview モジュール学習管理 - シートI/O・マイグレーション
 * @description module_control シートの読み書き、レイアウト管理、旧シート移行を担当します。
 */

/**
 * module_control から指定年度の計画行を抽出
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>=} allCycleRows - 事前取得済みの計画行全件
 * @param {Object=} layout - 事前取得済みレイアウト
 * @return {Array<Array<*>>} 行データ
 */
function readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear, allCycleRows, layout) {
  const rows = Array.isArray(allCycleRows) ? allCycleRows : readAllCyclePlanRows(controlSheet, layout);
  return rows.filter(function(row) {
    return Number(row[0]) === Number(fiscalYear);
  });
}

/**
 * 計画行を全件取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Object=} layout - 事前取得済みレイアウト
 * @return {Array<Array<*>>} 行データ
 */
function readAllCyclePlanRows(controlSheet, layout) {
  const sectionLayout = layout || getModuleControlLayout(controlSheet);
  const rowCount = sectionLayout.exceptionsMarkerRow - sectionLayout.planDataStartRow;
  if (rowCount <= 0) {
    return [];
  }

  const values = controlSheet.getRange(sectionLayout.planDataStartRow, 1, rowCount, MODULE_CONTROL_PLAN_HEADERS.length).getValues();
  return values.filter(function(row) {
    return row.some(function(value) {
      return isNonEmptyCell(value);
    });
  });
}

/**
 * モジュール管理用シートを初期化
 * 旧マルチシート構成から単一 module_control シートへ統合済みのため、
 * 戻り値の全プロパティは同一の controlSheet を指します（後方互換）。
 * @return {{controlSheet: GoogleAppsScript.Spreadsheet.Sheet, settingsSheet: GoogleAppsScript.Spreadsheet.Sheet, cyclePlanSheet: GoogleAppsScript.Spreadsheet.Sheet, dailyPlanSheet: GoogleAppsScript.Spreadsheet.Sheet, planSheet: GoogleAppsScript.Spreadsheet.Sheet, exceptionsSheet: GoogleAppsScript.Spreadsheet.Sheet, summarySheet: GoogleAppsScript.Spreadsheet.Sheet}}
 */
function initializeModuleHoursSheetsIfNeeded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = getOrCreateSheetByName(ss, MODULE_SHEET_NAMES.CONTROL);

  ensureModuleControlSheetLayout(controlSheet);
  ensureModuleSettingKeys();
  migrateLegacyModuleSheetsToControlIfNeeded(ss, controlSheet);
  ensureModuleControlSheetLayout(controlSheet);
  hideLegacyModuleSheets(ss);
  hideModuleControlSheetIfPossible(ss, controlSheet);

  return {
    controlSheet: controlSheet,
    settingsSheet: controlSheet,
    cyclePlanSheet: controlSheet,
    dailyPlanSheet: controlSheet,
    planSheet: controlSheet,
    exceptionsSheet: controlSheet,
    summarySheet: controlSheet
  };
}

/**
 * シート取得または作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シート
 */
function getOrCreateSheetByName(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * module_control レイアウトを保証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 */
function ensureModuleControlSheetLayout(controlSheet) {
  controlSheet.getRange(MODULE_CONTROL_DEFAULT_LAYOUT.VERSION_ROW, 1, 1, 2)
    .setValues([['MODULE_CONTROL_VERSION', MODULE_DATA_VERSION]]);

  const layout = getModuleControlLayout(controlSheet);

  controlSheet.getRange(layout.planMarkerRow, 1).setValue(MODULE_CONTROL_MARKERS.PLAN);
  controlSheet.getRange(layout.planHeaderRow, 1, 1, MODULE_CONTROL_PLAN_HEADERS.length)
    .setValues([MODULE_CONTROL_PLAN_HEADERS]);

  controlSheet.getRange(layout.exceptionsMarkerRow, 1).setValue(MODULE_CONTROL_MARKERS.EXCEPTIONS);
  controlSheet.getRange(layout.exceptionsHeaderRow, 1, 1, MODULE_CONTROL_EXCEPTION_HEADERS.length)
    .setValues([MODULE_CONTROL_EXCEPTION_HEADERS]);
}

/**
 * module_control のセクション位置を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @return {Object} レイアウト情報
 */
function getModuleControlLayout(controlSheet) {
  // 1回のシート読み取りで両マーカーを同時検索（冗長な findMarkerRow 呼び出しを排除）
  const maxRows = Math.max(controlSheet.getLastRow(), 200);
  const values = controlSheet.getRange(1, 1, maxRows, 1).getDisplayValues();

  let planMarkerRow = -1;
  let exceptionsMarkerRow = -1;

  for (let i = 0; i < values.length; i++) {
    const cellValue = String(values[i][0] || '').trim();
    if (cellValue === MODULE_CONTROL_MARKERS.PLAN && planMarkerRow < 0) {
      planMarkerRow = i + 1;
    }
    if (cellValue === MODULE_CONTROL_MARKERS.EXCEPTIONS) {
      exceptionsMarkerRow = i + 1;
    }
  }

  if (planMarkerRow < 1) {
    planMarkerRow = MODULE_CONTROL_DEFAULT_LAYOUT.PLAN_MARKER_ROW;
    controlSheet.getRange(planMarkerRow, 1).setValue(MODULE_CONTROL_MARKERS.PLAN);
  }

  if (exceptionsMarkerRow < 1 || exceptionsMarkerRow <= planMarkerRow + 2) {
    exceptionsMarkerRow = Math.max(
      MODULE_CONTROL_DEFAULT_LAYOUT.EXCEPTIONS_MARKER_ROW,
      planMarkerRow + 20,
      controlSheet.getLastRow() + 2
    );
    controlSheet.getRange(exceptionsMarkerRow, 1).setValue(MODULE_CONTROL_MARKERS.EXCEPTIONS);
  }

  return {
    planMarkerRow: planMarkerRow,
    planHeaderRow: planMarkerRow + 1,
    planDataStartRow: planMarkerRow + 2,
    exceptionsMarkerRow: exceptionsMarkerRow,
    exceptionsHeaderRow: exceptionsMarkerRow + 1,
    exceptionsDataStartRow: exceptionsMarkerRow + 2
  };
}

/**
 * 指定マーカー行を検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} marker - マーカー文字列
 * @param {boolean} useLast - 末尾一致を採用する場合 true
 * @return {number} 行番号（見つからない場合 -1）
 */
function findMarkerRow(sheet, marker, useLast) {
  const maxRows = Math.max(sheet.getLastRow(), 200);
  const values = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  let found = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === marker) {
      found = i + 1;
      if (!useLast) {
        break;
      }
    }
  }

  return found;
}

/**
 * 計画行を例外セクション直前へ追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Array<Array<*>>} rows - 追加行
 */
function appendCyclePlanRows(controlSheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const layout = getModuleControlLayout(controlSheet);
  const insertRow = layout.exceptionsMarkerRow;

  controlSheet.insertRowsBefore(insertRow, rows.length);
  controlSheet.getRange(insertRow, 1, rows.length, MODULE_CONTROL_PLAN_HEADERS.length).setValues(rows);
}

/**
 * 指定年度の計画行を置換
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>} replacementRows - 置換行
 */
function replaceCyclePlanRowsForFiscalYearInControl(controlSheet, fiscalYear, replacementRows) {
  const targetFiscalYear = Number(fiscalYear);
  const keptRows = readAllCyclePlanRows(controlSheet).filter(function(row) {
    return Number(row[0]) !== targetFiscalYear;
  });

  const mergedRows = keptRows.concat(replacementRows).sort(function(a, b) {
    if (Number(a[0]) !== Number(b[0])) {
      return Number(a[0]) - Number(b[0]);
    }
    return Number(a[1]) - Number(b[1]);
  });

  let layout = getModuleControlLayout(controlSheet);
  const currentCapacity = Math.max(layout.exceptionsMarkerRow - layout.planDataStartRow, 0);
  if (mergedRows.length > currentCapacity) {
    controlSheet.insertRowsBefore(layout.exceptionsMarkerRow, mergedRows.length - currentCapacity);
    layout = getModuleControlLayout(controlSheet);
  }

  const clearRowCount = Math.max(layout.exceptionsMarkerRow - layout.planDataStartRow, 0);
  if (clearRowCount > 0) {
    controlSheet.getRange(layout.planDataStartRow, 1, clearRowCount, MODULE_CONTROL_PLAN_HEADERS.length).clearContent();
  }

  if (mergedRows.length > 0) {
    controlSheet.getRange(layout.planDataStartRow, 1, mergedRows.length, MODULE_CONTROL_PLAN_HEADERS.length).setValues(mergedRows);
  }
}

/**
 * 例外行を末尾へ追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Array<Array<*>>} rows - 追加行
 */
function appendExceptionRows(controlSheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const layout = getModuleControlLayout(controlSheet);
  const start = findFirstEmptyExceptionRow(controlSheet, layout);
  controlSheet.getRange(start, 1, rows.length, MODULE_CONTROL_EXCEPTION_HEADERS.length).setValues(rows);
}

/**
 * 例外入力の最初の空行を返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Object} layout - レイアウト
 * @return {number} 行番号
 */
function findFirstEmptyExceptionRow(controlSheet, layout) {
  const lastRow = controlSheet.getLastRow();
  if (lastRow < layout.exceptionsDataStartRow) {
    return layout.exceptionsDataStartRow;
  }

  const values = controlSheet
    .getRange(layout.exceptionsDataStartRow, 1, lastRow - layout.exceptionsDataStartRow + 1, MODULE_CONTROL_EXCEPTION_HEADERS.length)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const empty = values[i].every(function(value) {
      return !isNonEmptyCell(value);
    });
    if (empty) {
      return layout.exceptionsDataStartRow + i;
    }
  }

  return lastRow + 1;
}

/**
 * 例外行を読み込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Object=} layout - 事前取得済みレイアウト
 * @return {Array<Object>} 例外行
 */
function readExceptionRows(controlSheet, layout) {
  const sectionLayout = layout || getModuleControlLayout(controlSheet);
  const lastRow = controlSheet.getLastRow();

  if (lastRow < sectionLayout.exceptionsDataStartRow) {
    return [];
  }

  const values = controlSheet
    .getRange(sectionLayout.exceptionsDataStartRow, 1, lastRow - sectionLayout.exceptionsDataStartRow + 1, MODULE_CONTROL_EXCEPTION_HEADERS.length)
    .getValues();

  const rows = [];
  values.forEach(function(row, index) {
    const hasValue = row.some(function(value) {
      return isNonEmptyCell(value);
    });

    if (!hasValue) {
      return;
    }

    rows.push({
      rowNumber: sectionLayout.exceptionsDataStartRow + index,
      date: row[0],
      grade: row[1],
      deltaSessions: row[2],
      reason: row[3],
      note: row[4]
    });
  });

  return rows;
}

/**
 * 年度別の計画行数をカウント
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 年度
 * @param {Array<Array<*>>=} cyclePlanRows - 事前取得済みの対象年度行
 * @return {number} 行数
 */
function countCyclePlanRowsForFiscalYear(controlSheet, fiscalYear, cyclePlanRows) {
  if (Array.isArray(cyclePlanRows)) {
    return cyclePlanRows.length;
  }
  return readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear).length;
}

/**
 * 年度別の例外行数をカウント
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 年度
 * @param {Array<Object>=} exceptionRows - 事前取得済み例外行
 * @return {number} 行数
 */
function countExceptionRowsForFiscalYear(controlSheet, fiscalYear, exceptionRows) {
  const rows = Array.isArray(exceptionRows) ? exceptionRows : readExceptionRows(controlSheet);
  return rows.filter(function(item) {
    const date = normalizeToDate(item.date);
    return !!date && getFiscalYear(date) === Number(fiscalYear);
  }).length;
}

/**
 * 旧モジュールシートから module_control へ移行
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 */
function migrateLegacyModuleSheetsToControlIfNeeded(ss, controlSheet) {
  const settings = readModuleSettingsMap();
  if (String(settings[MODULE_SETTING_KEYS.DATA_VERSION] || '').trim() === MODULE_DATA_VERSION) {
    return;
  }

  if (readAllCyclePlanRows(controlSheet).length === 0) {
    const legacyCycleSheet = ss.getSheetByName(MODULE_SHEET_NAMES.CYCLE_PLAN);
    if (legacyCycleSheet && legacyCycleSheet.getLastRow() > 1) {
      const colCount = Math.min(legacyCycleSheet.getLastColumn(), MODULE_CONTROL_PLAN_HEADERS.length);
      const values = legacyCycleSheet.getRange(2, 1, legacyCycleSheet.getLastRow() - 1, colCount).getValues();

      const rows = values.map(function(row) {
        const padded = new Array(MODULE_CONTROL_PLAN_HEADERS.length).fill('');
        for (let i = 0; i < colCount; i++) {
          padded[i] = row[i];
        }
        return padded;
      }).filter(function(row) {
        return row.some(function(value) {
          return isNonEmptyCell(value);
        });
      });

      appendCyclePlanRows(controlSheet, rows);
    }
  }

  if (readExceptionRows(controlSheet).length === 0) {
    const legacyExceptionSheet = ss.getSheetByName(MODULE_SHEET_NAMES.EXCEPTIONS);
    if (legacyExceptionSheet && legacyExceptionSheet.getLastRow() > 1) {
      const colCount = Math.min(legacyExceptionSheet.getLastColumn(), MODULE_CONTROL_EXCEPTION_HEADERS.length);
      const headerCol3 = String(legacyExceptionSheet.getRange(1, 3).getValue() || '').trim();
      const values = legacyExceptionSheet.getRange(2, 1, legacyExceptionSheet.getLastRow() - 1, colCount).getValues();

      const rows = values.map(function(row) {
        const padded = new Array(MODULE_CONTROL_EXCEPTION_HEADERS.length).fill('');
        for (let i = 0; i < colCount; i++) {
          padded[i] = row[i];
        }

        if (headerCol3 === 'delta_units') {
          padded[2] = toNumberOrZero(padded[2]) * 3;
        }

        return padded;
      }).filter(function(row) {
        return row.some(function(value) {
          return isNonEmptyCell(value);
        });
      });

      appendExceptionRows(controlSheet, rows);
    }
  }

  const legacySettingsSheet = ss.getSheetByName(MODULE_SHEET_NAMES.SETTINGS);
  if (legacySettingsSheet) {
    migrateLegacySettingsFromSheet(legacySettingsSheet);
  }

  upsertModuleSettingsValues(null, {
    DATA_VERSION: MODULE_DATA_VERSION
  });
}

/**
 * 旧設定シートからプロパティへ移行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - module_settings
 */
function migrateLegacySettingsFromSheet(settingsSheet) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const values = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const legacyMap = {};

  values.forEach(function(row) {
    if (isNonEmptyCell(row[0])) {
      legacyMap[String(row[0])] = row[1];
    }
  });

  const current = readModuleSettingsMap();
  const updates = {};

  Object.keys(MODULE_SETTING_KEYS).forEach(function(keyName) {
    const key = MODULE_SETTING_KEYS[keyName];
    if (!isNonEmptyCell(current[key]) && Object.prototype.hasOwnProperty.call(legacyMap, key)) {
      updates[key] = legacyMap[key];
    }
  });

  if (Object.keys(updates).length > 0) {
    upsertModuleSettingsValues(null, updates);
  }
}

/**
 * 旧シートを非表示化（1画面運用）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 */
function hideLegacyModuleSheets(ss) {
  const legacyNames = [
    MODULE_SHEET_NAMES.SETTINGS,
    MODULE_SHEET_NAMES.CYCLE_PLAN,
    MODULE_SHEET_NAMES.DAILY_PLAN,
    MODULE_SHEET_NAMES.PLAN,
    MODULE_SHEET_NAMES.EXCEPTIONS,
    MODULE_SHEET_NAMES.SUMMARY
  ];

  const active = ss.getActiveSheet();
  const activeName = active ? active.getName() : '';

  legacyNames.forEach(function(name) {
    if (name === MODULE_SHEET_NAMES.CONTROL) {
      return;
    }

    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      return;
    }
    if (sheet.getName() === activeName) {
      return;
    }

    try {
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    } catch (error) {
      Logger.log('[WARNING] 旧シート非表示に失敗: ' + name + ' / ' + error.toString());
    }
  });
}

/**
 * module_control を必要時のみ表示するため通常は非表示化
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 */
function hideModuleControlSheetIfPossible(ss, controlSheet) {
  try {
    const activeSheet = ss.getActiveSheet();
    if (activeSheet && activeSheet.getSheetId() === controlSheet.getSheetId()) {
      const fallbackSheet = findFallbackSheetForHiding(ss, controlSheet.getSheetId());
      if (!fallbackSheet) {
        return;
      }
      if (fallbackSheet.isSheetHidden()) {
        fallbackSheet.showSheet();
      }
      ss.setActiveSheet(fallbackSheet);
    }

    if (!controlSheet.isSheetHidden()) {
      controlSheet.hideSheet();
    }
  } catch (error) {
    Logger.log('[WARNING] module_control 非表示に失敗: ' + error.toString());
  }
}

/**
 * 非表示対象以外で切替可能なシートを取得
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {number} excludedSheetId - 除外対象シートID
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} 切替先シート
 */
function findFallbackSheetForHiding(ss, excludedSheetId) {
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (sheet.getSheetId() !== excludedSheetId && !sheet.isSheetHidden()) {
      return sheet;
    }
  }

  for (let j = 0; j < sheets.length; j++) {
    const hiddenSheet = sheets[j];
    if (hiddenSheet.getSheetId() !== excludedSheetId) {
      return hiddenSheet;
    }
  }

  return null;
}

/**
 * module_settings の必須キーを保証（プロパティ）
 */
function ensureModuleSettingKeys() {
  const requiredKeys = [
    MODULE_SETTING_KEYS.PLAN_START_DATE,
    MODULE_SETTING_KEYS.PLAN_END_DATE,
    MODULE_SETTING_KEYS.LAST_GENERATED_AT,
    MODULE_SETTING_KEYS.LAST_DAILY_PLAN_COUNT,
    MODULE_SETTING_KEYS.DATA_VERSION
  ];

  const map = readModuleSettingsMap();
  const updates = {};

  requiredKeys.forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(map, key)) {
      updates[key] = '';
    }
  });

  if (Object.keys(updates).length > 0) {
    upsertModuleSettingsValues(null, updates);
  }
}

/**
 * module settings を key-value マップ化（プロパティ）
 * @return {Object} 設定マップ
 */
function readModuleSettingsMap() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  const map = {};

  Object.keys(props).forEach(function(rawKey) {
    if (rawKey.indexOf(MODULE_SETTINGS_PREFIX) !== 0) {
      return;
    }
    const key = rawKey.substring(MODULE_SETTINGS_PREFIX.length);
    map[key] = props[rawKey];
  });

  return map;
}

/**
 * module settings を更新または追加（プロパティ）
 * @param {*} settingsSheet - 旧互換引数（未使用）
 * @param {Object} updates - 追加/更新値
 */
function upsertModuleSettingsValues(settingsSheet, updates) {
  const docProps = PropertiesService.getDocumentProperties();
  const serialized = {};

  Object.keys(updates).forEach(function(key) {
    serialized[MODULE_SETTINGS_PREFIX + key] = serializeModuleSettingValue(key, updates[key]);
  });

  docProps.setProperties(serialized, false);
}

/**
 * 設定値を文字列にシリアライズ
 * @param {string} key - 設定キー
 * @param {*} value - 値
 * @return {string} 文字列値
 */
function serializeModuleSettingValue(key, value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    if (key === MODULE_SETTING_KEYS.LAST_GENERATED_AT) {
      return value.toISOString();
    }
    return formatInputDate(value);
  }

  return String(value);
}

/**
 * fiscal_year キーで対象年度行を置換（汎用ユーティリティ）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Array<*>>} rows - 書き込み行
 * @param {number} fiscalYear - 対象年度
 * @param {number} fiscalYearColumnIndex - fiscal_year列index(0-based)
 * @param {number} columnCount - 列数
 */
function replaceRowsForFiscalYear(sheet, rows, fiscalYear, fiscalYearColumnIndex, columnCount) {
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, columnCount).getValues() : [];
  const targetFiscalYear = Number(fiscalYear);

  const kept = existing.filter(function(row) {
    const rawFiscalYear = row[fiscalYearColumnIndex];
    const rowFiscalYear = Number(rawFiscalYear);

    if (Number.isFinite(rowFiscalYear)) {
      return rowFiscalYear !== targetFiscalYear;
    }

    const text = String(rawFiscalYear === null || rawFiscalYear === undefined ? '' : rawFiscalYear).trim();
    const legacyMatch = text.match(/^(\d{4})(?:[-\/].*)?$/);
    if (legacyMatch) {
      return Number(legacyMatch[1]) !== targetFiscalYear;
    }

    return true;
  });

  const merged = kept.concat(rows);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, columnCount).clearContent();
  }

  if (merged.length > 0) {
    sheet.getRange(2, 1, merged.length, columnCount).setValues(merged);
  }
}

/**
 * 指定年度の行数をカウント（汎用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} fiscalYear - 年度
 * @param {number} fiscalYearColumnIndex - fiscal_year列index(0-based)
 * @return {number} 行数
 */
function countRowsByFiscalYear(sheet, fiscalYear, fiscalYearColumnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, fiscalYearColumnIndex + 1).getValues();
  let count = 0;

  values.forEach(function(row) {
    if (Number(row[fiscalYearColumnIndex]) === Number(fiscalYear)) {
      count++;
    }
  });

  return count;
}
