/**
 * @fileoverview モジュール学習管理 - シートI/O・マイグレーション
 * @description module_control シートの読み書き、レイアウト管理、旧シート移行を担当します。
 */

// ── Per-execution caches ──
// GAS の各実行（メニュー操作やトリガー）内で同じデータを繰り返し読まないよう、
// 初回取得結果をキャッシュする。実行終了時に自動破棄される。
let moduleHoursSheetsCache_ = null;
let moduleSettingsMapCache_ = null;
let moduleControlLayoutCache_ = null;

/**
 * module_control レイアウトキャッシュを無効化
 * 行挿入などでセクション境界が変わった直後に呼び出す。
 */
function invalidateModuleControlLayoutCache_() {
  moduleControlLayoutCache_ = null;
}

/**
 * module_control から指定年度の年間計画時数行を抽出
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>=} allRows - 事前取得済みの計画行全件
 * @param {Object=} layout - 事前取得済みレイアウト
 * @return {Array<Array<*>>} 行データ
 */
function readAnnualTargetRowsByFiscalYear(controlSheet, fiscalYear, allRows, layout) {
  const rows = Array.isArray(allRows) ? allRows : readAllAnnualTargetRows(controlSheet, layout);
  return rows.filter(function(row) {
    return Number(row[0]) === Number(fiscalYear);
  });
}

/**
 * 年間計画時数行を全件取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Object=} layout - 事前取得済みレイアウト
 * @return {Array<Array<*>>} 行データ
 */
function readAllAnnualTargetRows(controlSheet, layout) {
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
 * 旧マルチシート構成から単一 module_control シートへ統合済み。
 * @return {GoogleAppsScript.Spreadsheet.Sheet} controlSheet
 */
function initializeModuleHoursSheetsIfNeeded() {
  if (moduleHoursSheetsCache_) {
    return moduleHoursSheetsCache_;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = getOrCreateSheetByName(ss, MODULE_SHEET_NAMES.CONTROL);

  ensureModuleSettingKeys();
  ensureDataVersionIsLatest();
  ensureModuleControlSheetLayout(controlSheet);
  hideModuleControlSheetIfPossible(ss, controlSheet);

  moduleHoursSheetsCache_ = controlSheet;
  return controlSheet;
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
  if (moduleControlLayoutCache_ &&
      moduleControlLayoutCache_.sheetId === controlSheet.getSheetId()) {
    return moduleControlLayoutCache_.layout;
  }

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

  const layout = {
    planMarkerRow: planMarkerRow,
    planHeaderRow: planMarkerRow + 1,
    planDataStartRow: planMarkerRow + 2,
    exceptionsMarkerRow: exceptionsMarkerRow,
    exceptionsHeaderRow: exceptionsMarkerRow + 1,
    exceptionsDataStartRow: exceptionsMarkerRow + 2
  };

  moduleControlLayoutCache_ = {
    sheetId: controlSheet.getSheetId(),
    layout: layout
  };
  return layout;
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
 * 年間計画時数行を例外セクション直前へ追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {Array<Array<*>>} rows - 追加行
 */
function appendAnnualTargetRows(controlSheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const layout = getModuleControlLayout(controlSheet);
  const insertRow = layout.exceptionsMarkerRow;

  controlSheet.insertRowsBefore(insertRow, rows.length);
  invalidateModuleControlLayoutCache_();
  controlSheet.getRange(insertRow, 1, rows.length, MODULE_CONTROL_PLAN_HEADERS.length).setValues(rows);
}

/**
 * 指定年度の年間計画時数行を置換
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>} replacementRows - 置換行
 */
function replaceAnnualTargetRowsForFiscalYearInControl(controlSheet, fiscalYear, replacementRows) {
  const targetFiscalYear = Number(fiscalYear);
  const keptRows = readAllAnnualTargetRows(controlSheet).filter(function(row) {
    return Number(row[0]) !== targetFiscalYear;
  });

  const mergedRows = keptRows.concat(replacementRows).sort(function(a, b) {
    return Number(a[0]) - Number(b[0]);
  });

  let layout = getModuleControlLayout(controlSheet);
  const currentCapacity = Math.max(layout.exceptionsMarkerRow - layout.planDataStartRow, 0);
  if (mergedRows.length > currentCapacity) {
    controlSheet.insertRowsBefore(layout.exceptionsMarkerRow, mergedRows.length - currentCapacity);
    invalidateModuleControlLayoutCache_();
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
 * 年度別の年間計画時数行数をカウント
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 年度
 * @param {Array<Array<*>>=} annualTargetRows - 事前取得済みの対象年度行
 * @return {number} 行数
 */
function countAnnualTargetRowsForFiscalYear(controlSheet, fiscalYear, annualTargetRows) {
  if (Array.isArray(annualTargetRows)) {
    return annualTargetRows.length;
  }
  return readAnnualTargetRowsByFiscalYear(controlSheet, fiscalYear).length;
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
 * データバージョンが最新でなければ更新する
 */
function ensureDataVersionIsLatest() {
  const settings = readModuleSettingsMap();
  const currentVersion = String(settings[MODULE_SETTING_KEYS.DATA_VERSION] || '').trim();

  if (currentVersion !== MODULE_DATA_VERSION) {
    upsertModuleSettingsValues({
      DATA_VERSION: MODULE_DATA_VERSION
    });
  }
}

/**
 * V4形式の計画行を構築
 * @param {number} fiscalYear - 年度
 * @param {number} grade - 学年
 * @param {string} mode - 'annual' or 'monthly'
 * @param {number} annualKoma - 年間計画時数
 * @param {Object|null} monthlyKoma - 月別計画時数（monthlyモード時）
 * @param {string=} note - メモ
 * @return {Array<*>} MODULE_CONTROL_PLAN_HEADERS形式の行
 */
function buildV4PlanRow(fiscalYear, grade, mode, annualKoma, monthlyKoma, note) {
  const row = new Array(MODULE_CONTROL_PLAN_HEADERS.length).fill('');
  row[0] = Number(fiscalYear);
  row[1] = Number(grade);
  row[2] = mode || MODULE_PLAN_MODE_ANNUAL;
  if (mode === MODULE_PLAN_MODE_MONTHLY && monthlyKoma) {
    [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].forEach(function(m, i) {
      row[3 + i] = toNumberOrZero(monthlyKoma[m]);
    });
  }
  row[15] = toNumberOrZero(annualKoma);
  row[16] = note || '';
  return row;
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

  return null;
}

/**
 * module_settings の必須キーを保証（プロパティ）
 */
function ensureModuleSettingKeys() {
  const requiredKeys = [
    MODULE_SETTING_KEYS.PLAN_START_DATE,
    MODULE_SETTING_KEYS.PLAN_END_DATE,
    MODULE_SETTING_KEYS.WEEKDAYS_ENABLED,
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
    upsertModuleSettingsValues(updates);
  }
}

/**
 * module settings を key-value マップ化（プロパティ）
 * @return {Object} 設定マップ
 */
function readModuleSettingsMap() {
  if (moduleSettingsMapCache_) {
    // 呼び出し元が返却値を変更してもキャッシュに影響しないよう浅いコピーを返す
    const copy = {};
    Object.keys(moduleSettingsMapCache_).forEach(function(k) {
      copy[k] = moduleSettingsMapCache_[k];
    });
    return copy;
  }

  const props = PropertiesService.getDocumentProperties().getProperties();
  const map = {};

  Object.keys(props).forEach(function(rawKey) {
    if (rawKey.indexOf(MODULE_SETTINGS_PREFIX) !== 0) {
      return;
    }
    const key = rawKey.substring(MODULE_SETTINGS_PREFIX.length);
    map[key] = props[rawKey];
  });

  moduleSettingsMapCache_ = map;

  // 浅いコピーを返す
  const result = {};
  Object.keys(map).forEach(function(k) {
    result[k] = map[k];
  });
  return result;
}

/**
 * module settings を更新または追加（プロパティ）
 * @param {Object} updates - 追加/更新値
 */
function upsertModuleSettingsValues(updates) {
  const docProps = PropertiesService.getDocumentProperties();
  const serialized = {};

  Object.keys(updates).forEach(function(key) {
    serialized[MODULE_SETTINGS_PREFIX + key] = serializeModuleSettingValue(key, updates[key]);
  });

  docProps.setProperties(serialized, false);

  // キャッシュを無効化して次回読み取り時に最新値を取得させる
  moduleSettingsMapCache_ = null;
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

  if (key === MODULE_SETTING_KEYS.WEEKDAYS_ENABLED && Array.isArray(value)) {
    return serializeWeekdays(value);
  }

  return String(value);
}

/**
 * 保存済みの実施曜日を取得（未設定時はデフォルト）
 * @param {Object=} settingsMap - 事前取得済み設定マップ
 * @return {Array<number>} 有効曜日配列（getDay()値）
 */
function getEnabledWeekdays(settingsMap) {
  const map = settingsMap || readModuleSettingsMap();
  const raw = map[MODULE_SETTING_KEYS.WEEKDAYS_ENABLED];

  if (!raw || String(raw).trim() === '') {
    return MODULE_DEFAULT_WEEKDAYS_ENABLED.slice();
  }

  const parsed = String(raw).split(',')
    .map(function(s) { return parseInt(s.trim(), 10); })
    .filter(function(n) { return Number.isInteger(n) && n >= 1 && n <= 5; });

  if (parsed.length === 0) {
    return MODULE_DEFAULT_WEEKDAYS_ENABLED.slice();
  }

  return parsed;
}

/**
 * 実施曜日をシリアライズ
 * @param {Array<number>} weekdays - 有効曜日配列
 * @return {string} カンマ区切り文字列
 */
function serializeWeekdays(weekdays) {
  if (!Array.isArray(weekdays) || weekdays.length === 0) {
    return MODULE_DEFAULT_WEEKDAYS_ENABLED.join(',');
  }
  const valid = weekdays
    .filter(function(n) { return Number.isInteger(n) && n >= 1 && n <= 5; })
    .sort(function(a, b) { return a - b; });
  if (valid.length === 0) {
    return MODULE_DEFAULT_WEEKDAYS_ENABLED.join(',');
  }
  return valid.join(',');
}
