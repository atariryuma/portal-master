/**
 * @fileoverview モジュール時数管理機能
 * @description 計画期間設定・学校日ベースの自動計画生成・例外反映・累計時数統合を提供します。
 */

/**
 * モジュール計画期間設定ダイアログを表示
 */
function showModulePlanningDialog() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('modulePlanningDialog')
      .setWidth(460)
      .setHeight(380);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'モジュール計画期間設定');
  } catch (error) {
    showAlert('モジュール計画ダイアログの表示に失敗しました: ' + error.toString(), 'エラー');
  }
}

/**
 * ダイアログ用の初期値を返却
 * @return {Object} 開始日・終了日（yyyy-MM-dd）
 */
function getModulePlanningDefaults() {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const fallbackDate = normalizeToDate(new Date());
  const range = getModulePlanningRangeFromSettings(sheets.settingsSheet, fallbackDate);
  return {
    startDate: formatInputDate(range.startDate),
    endDate: formatInputDate(range.endDate)
  };
}

/**
 * モジュール計画期間を保存し、計画を再生成
 * @param {string|Date} startDate - 開始日
 * @param {string|Date} endDate - 終了日
 * @return {string} 完了メッセージ
 */
function saveModulePlanningRange(startDate, endDate) {
  const start = normalizeToDate(startDate);
  const end = normalizeToDate(endDate);

  if (!start || !end) {
    throw new Error('開始日・終了日は yyyy-MM-dd 形式で入力してください。');
  }
  if (start > end) {
    throw new Error('開始日は終了日以前の日付を指定してください。');
  }

  const result = rebuildModulePlanFromRange(start, end);
  return [
    'モジュール計画を更新しました。',
    '期間: ' + formatInputDate(result.startDate) + ' ～ ' + formatInputDate(result.endDate),
    '生成件数: ' + result.recordCount + '件'
  ].join('\n');
}

/**
 * 指定期間で module_plan を再生成
 * @param {string|Date} startDate - 開始日
 * @param {string|Date} endDate - 終了日
 * @return {Object} 再生成結果
 */
function rebuildModulePlanFromRange(startDate, endDate) {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const start = normalizeToDate(startDate);
  const end = normalizeToDate(endDate);

  if (!start || !end) {
    throw new Error('計画期間の日付が不正です。');
  }
  if (start > end) {
    throw new Error('計画期間の開始日と終了日の順序が不正です。');
  }

  const planMap = buildSchoolDayPlanMap(start, end);
  const generatedAt = new Date();
  const recordCount = writeModulePlanSheet(sheets.planSheet, planMap, generatedAt);

  upsertModuleSettingsValues(sheets.settingsSheet, {
    PLAN_START_DATE: start,
    PLAN_END_DATE: end,
    LAST_GENERATED_AT: generatedAt
  });

  Logger.log('[INFO] module_plan を再生成しました（件数: ' + recordCount + '）');
  return {
    startDate: start,
    endDate: end,
    generatedAt: generatedAt,
    recordCount: recordCount,
    planMap: planMap
  };
}

/**
 * モジュール時数を集計し、累計時数シートへ統合出力
 * @param {Date|string} baseDate - 集計基準日
 * @return {Object} 集計結果
 */
function syncModuleHoursWithCumulative(baseDate) {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const normalizedBaseDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const range = getModulePlanningRangeFromSettings(sheets.settingsSheet, normalizedBaseDate);

  rebuildModulePlanFromRange(range.startDate, range.endDate);

  let summaryBasePlanMap;
  if (normalizedBaseDate < range.startDate) {
    // 計画開始日前は0件として扱う
    summaryBasePlanMap = createEmptyPlanMap(range.startDate, range.startDate);
  } else {
    const effectiveEndDate = normalizedBaseDate > range.endDate ? range.endDate : normalizedBaseDate;
    summaryBasePlanMap = buildSchoolDayPlanMap(range.startDate, effectiveEndDate);
  }

  const planMap = applyModuleExceptions(summaryBasePlanMap, normalizedBaseDate);

  writeModuleSummary(planMap, normalizedBaseDate, sheets.summarySheet);
  writeModuleToCumulativeSheet(planMap, normalizedBaseDate);

  Logger.log('[INFO] モジュール時数を累計時数へ統合しました（基準日: ' + formatInputDate(normalizedBaseDate) + '）');

  return {
    baseDate: normalizedBaseDate,
    startDate: range.startDate,
    endDate: range.endDate
  };
}

/**
 * 学校日判定に基づいて月別・学年別の計画マップを構築
 * @param {Date|string} startDate - 開始日
 * @param {Date|string} endDate - 終了日
 * @return {Object} 計画マップ
 */
function buildSchoolDayPlanMap(startDate, endDate) {
  const start = normalizeToDate(startDate);
  const end = normalizeToDate(endDate);

  if (!start || !end) {
    throw new Error('学校日計画の期間指定が不正です。');
  }
  if (start > end) {
    throw new Error('学校日計画の開始日と終了日の順序が不正です。');
  }

  const planMap = createEmptyPlanMap(start, end);
  const scheduleSheet = getAnnualScheduleSheetOrThrow();
  const data = scheduleSheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const uniqueDateGrade = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);

    if (!date) {
      continue;
    }
    if (date < start || date > end) {
      continue;
    }

    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      continue;
    }

    const grade = Number(row[SCHEDULE_COLUMNS.GRADE]);
    if (!Number.isInteger(grade) || grade < 1 || grade > 6) {
      continue;
    }

    let hasSchoolData = false;
    for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
      if (isNonEmptyCell(row[j])) {
        hasSchoolData = true;
        break;
      }
    }
    if (!hasSchoolData) {
      continue;
    }

    const dateKey = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
    const uniqueKey = dateKey + '_' + grade;
    if (uniqueDateGrade[uniqueKey]) {
      continue;
    }
    uniqueDateGrade[uniqueKey] = true;

    const monthKey = Utilities.formatDate(date, tz, 'yyyy-MM');
    const gradeEntry = planMap.byMonth[monthKey] && planMap.byMonth[monthKey][grade];
    if (!gradeEntry) {
      continue;
    }

    gradeEntry.planned_units += 1;
    gradeEntry.school_days_count += 1;
  }

  return planMap;
}

/**
 * module_exceptions の差分を計画マップへ反映
 * @param {Object} planMap - 計画マップ
 * @param {Date|string} baseDate - 集計基準日
 * @return {Object} 差分反映後の計画マップ
 */
function applyModuleExceptions(planMap, baseDate) {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const exceptionSheet = sheets.exceptionsSheet;
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());

  // 初期化
  Object.keys(planMap.byMonth).forEach(function(monthKey) {
    for (let grade = 1; grade <= 6; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      entry.delta_units = 0;
      entry.actual_units = entry.planned_units;
      entry.diff_units = 0;
    }
  });

  const lastRow = exceptionSheet.getLastRow();
  if (lastRow > 1) {
    const values = exceptionSheet.getRange(2, 1, lastRow - 1, 5).getValues();

    values.forEach(function(row, index) {
      const exceptionDate = normalizeToDate(row[0]);
      const grade = Number(row[1]);
      const delta = Number(row[2]);

      if (!exceptionDate || !Number.isInteger(grade) || grade < 1 || grade > 6 || isNaN(delta)) {
        Logger.log('[WARNING] module_exceptions の入力不正をスキップしました（行: ' + (index + 2) + '）');
        return;
      }

      if (exceptionDate > cutoffDate) {
        return;
      }

      const monthKey = formatMonthKey(exceptionDate);
      if (!planMap.byMonth[monthKey] || !planMap.byMonth[monthKey][grade]) {
        Logger.log('[WARNING] 計画範囲外の例外をスキップしました（行: ' + (index + 2) + '）');
        return;
      }

      planMap.byMonth[monthKey][grade].delta_units += delta;
    });
  }

  Object.keys(planMap.byMonth).forEach(function(monthKey) {
    for (let grade = 1; grade <= 6; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      entry.actual_units = Math.max(entry.planned_units + entry.delta_units, 0);
      entry.diff_units = entry.actual_units - entry.planned_units;
    }
  });

  return planMap;
}

/**
 * モジュール管理用シートを初期化
 * @return {Object} シート参照
 */
function initializeModuleHoursSheetsIfNeeded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const definitions = [
    { name: MODULE_SHEET_NAMES.SETTINGS, headers: ['key', 'value'] },
    { name: MODULE_SHEET_NAMES.PLAN, headers: ['year_month', 'grade', 'planned_units', 'school_days_count', 'generated_at'] },
    { name: MODULE_SHEET_NAMES.EXCEPTIONS, headers: ['date', 'grade', 'delta_units', 'reason', 'note'] },
    { name: MODULE_SHEET_NAMES.SUMMARY, headers: ['fiscal_year', 'year_month', 'grade', 'planned_units', 'delta_units', 'actual_units', 'diff_units', 'calculated_at'] }
  ];

  const sheets = {};

  definitions.forEach(function(definition) {
    let sheet = ss.getSheetByName(definition.name);
    if (!sheet) {
      sheet = ss.insertSheet(definition.name);
    }
    ensureModuleSheetHeaders(sheet, definition.headers);
    sheets[definition.name] = sheet;
  });

  ensureModuleSettingKeys(sheets[MODULE_SHEET_NAMES.SETTINGS]);
  return {
    settingsSheet: sheets[MODULE_SHEET_NAMES.SETTINGS],
    planSheet: sheets[MODULE_SHEET_NAMES.PLAN],
    exceptionsSheet: sheets[MODULE_SHEET_NAMES.EXCEPTIONS],
    summarySheet: sheets[MODULE_SHEET_NAMES.SUMMARY]
  };
}

/**
 * シートヘッダーを保証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<string>} headers - ヘッダー配列
 */
function ensureModuleSheetHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsUpdate = headers.some(function(header, index) {
    return String(current[index] || '').trim() !== header;
  });

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * module_settings の必須キーを保証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - module_settings
 */
function ensureModuleSettingKeys(settingsSheet) {
  const requiredKeys = [
    MODULE_SETTING_KEYS.PLAN_START_DATE,
    MODULE_SETTING_KEYS.PLAN_END_DATE,
    MODULE_SETTING_KEYS.LAST_GENERATED_AT
  ];
  const map = readModuleSettingsMap(settingsSheet);
  const updates = {};

  requiredKeys.forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(map, key)) {
      updates[key] = '';
    }
  });

  if (Object.keys(updates).length > 0) {
    upsertModuleSettingsValues(settingsSheet, updates);
  }
}

/**
 * module_settings を key-value マップ化
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - module_settings
 * @return {Object} 設定マップ
 */
function readModuleSettingsMap(settingsSheet) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }

  const values = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};

  values.forEach(function(row) {
    const key = row[0];
    if (key) {
      map[String(key)] = row[1];
    }
  });

  return map;
}

/**
 * module_settings のキーを更新または追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - module_settings
 * @param {Object} updates - 追加/更新値
 */
function upsertModuleSettingsValues(settingsSheet, updates) {
  const lastRow = settingsSheet.getLastRow();
  const values = lastRow > 1 ? settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const keyRowMap = {};

  values.forEach(function(row, index) {
    if (row[0]) {
      keyRowMap[String(row[0])] = index + 2;
    }
  });

  Object.keys(updates).forEach(function(key) {
    let rowNumber = keyRowMap[key];
    if (!rowNumber) {
      rowNumber = settingsSheet.getLastRow() + 1;
    }
    settingsSheet.getRange(rowNumber, 1, 1, 2).setValues([[key, updates[key]]]);
  });
}

/**
 * 保存済み期間を取得（未設定時は当該年度のデフォルト期間）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - module_settings
 * @param {Date} fallbackDate - 未設定時の基準日
 * @return {Object} 計画期間
 */
function getModulePlanningRangeFromSettings(settingsSheet, fallbackDate) {
  const map = readModuleSettingsMap(settingsSheet);
  const start = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_START_DATE]);
  const end = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_END_DATE]);

  if (start && end && start <= end) {
    return { startDate: start, endDate: end };
  }

  const defaultRange = getDefaultModulePlanningRange(fallbackDate);
  upsertModuleSettingsValues(settingsSheet, {
    PLAN_START_DATE: defaultRange.startDate,
    PLAN_END_DATE: defaultRange.endDate
  });

  return defaultRange;
}

/**
 * 指定日を含む年度のデフォルト期間（4月1日～翌3月31日）
 * @param {Date} baseDate - 基準日
 * @return {Object} 期間
 */
function getDefaultModulePlanningRange(baseDate) {
  const date = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const fiscalYear = getFiscalYear(date);
  const startDate = new Date(fiscalYear, MODULE_FISCAL_YEAR_START_MONTH - 1, 1);
  const endDate = new Date(fiscalYear + 1, MODULE_FISCAL_YEAR_START_MONTH - 1, 0);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);
  return { startDate: startDate, endDate: endDate };
}

/**
 * module_plan を書き込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} planSheet - module_plan
 * @param {Object} planMap - 計画マップ
 * @param {Date} generatedAt - 生成日時
 * @return {number} 書き込み件数
 */
function writeModulePlanSheet(planSheet, planMap, generatedAt) {
  const existingRows = planSheet.getLastRow() - 1;
  if (existingRows > 0) {
    planSheet.getRange(2, 1, existingRows, 5).clearContent();
  }

  const rows = [];
  const monthKeys = Object.keys(planMap.byMonth).sort();

  monthKeys.forEach(function(monthKey) {
    for (let grade = 1; grade <= 6; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      rows.push([monthKey, grade, entry.planned_units, entry.school_days_count, generatedAt]);
    }
  });

  if (rows.length > 0) {
    planSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  return rows.length;
}

/**
 * module_summary を書き込み
 * @param {Object} planMap - 計画マップ
 * @param {Date} baseDate - 集計基準日
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet - module_summary
 */
function writeModuleSummary(planMap, baseDate, summarySheet) {
  const existingRows = summarySheet.getLastRow() - 1;
  if (existingRows > 0) {
    summarySheet.getRange(2, 1, existingRows, 8).clearContent();
  }

  const cutoffMonthKey = formatMonthKey(baseDate);
  const calculatedAt = new Date();
  const rows = [];
  const monthKeys = Object.keys(planMap.byMonth).sort();

  monthKeys.forEach(function(monthKey) {
    if (monthKeyCompare(monthKey, cutoffMonthKey) > 0) {
      return;
    }

    const fiscalYear = getFiscalYearFromMonthKey(monthKey);
    for (let grade = 1; grade <= 6; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      rows.push([
        fiscalYear,
        monthKey,
        grade,
        entry.planned_units,
        entry.delta_units,
        entry.actual_units,
        entry.diff_units,
        calculatedAt
      ]);
    }
  });

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }
}

/**
 * 累計時数シートへモジュール累計を出力
 * @param {Object} planMap - 計画マップ
 * @param {Date} baseDate - 集計基準日
 */
function writeModuleToCumulativeSheet(planMap, baseDate) {
  const cumulativeSheet = getSheetByNameOrThrow('累計時数');
  const totalsByGrade = buildYtdTotals(planMap, baseDate);

  cumulativeSheet
    .getRange(2, MODULE_CUMULATIVE_COLUMNS.PLAN, 1, 3)
    .setValues([['MOD計画累計', 'MOD実施累計', 'MOD差分']]);

  const rows = [];
  for (let grade = 1; grade <= 6; grade++) {
    const total = totalsByGrade[grade];
    rows.push([total.planned, total.actual, total.diff]);
  }

  cumulativeSheet.getRange(3, MODULE_CUMULATIVE_COLUMNS.PLAN, rows.length, 3).setValues(rows);
}

/**
 * 年度累計（4月開始）を学年別に算出
 * @param {Object} planMap - 計画マップ
 * @param {Date} baseDate - 集計基準日
 * @return {Object} 学年別累計
 */
function buildYtdTotals(planMap, baseDate) {
  const fiscalYear = getFiscalYear(baseDate);
  const cutoffMonthKey = formatMonthKey(baseDate);
  const startMonthKey = fiscalYear + '-' + String(MODULE_FISCAL_YEAR_START_MONTH).padStart(2, '0');
  const totals = {};

  for (let grade = 1; grade <= 6; grade++) {
    totals[grade] = { planned: 0, actual: 0, diff: 0 };
  }

  Object.keys(planMap.byMonth).sort().forEach(function(monthKey) {
    if (monthKeyCompare(monthKey, startMonthKey) < 0 || monthKeyCompare(monthKey, cutoffMonthKey) > 0) {
      return;
    }

    for (let grade = 1; grade <= 6; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      totals[grade].planned += entry.planned_units;
      totals[grade].actual += entry.actual_units;
      totals[grade].diff += entry.diff_units;
    }
  });

  return totals;
}

/**
 * 年度（4月開始）を取得
 * @param {Date|string} date - 対象日
 * @param {number} startMonth - 年度開始月
 * @return {number} 年度
 */
function getFiscalYear(date, startMonth) {
  const targetDate = normalizeToDate(date);
  const start = startMonth || MODULE_FISCAL_YEAR_START_MONTH;

  if (!targetDate) {
    throw new Error('年度計算対象の日付が不正です。');
  }

  const month = targetDate.getMonth() + 1;
  return month >= start ? targetDate.getFullYear() : targetDate.getFullYear() - 1;
}

/**
 * monthKey(yyyy-MM) から年度を取得
 * @param {string} monthKey - 月キー
 * @return {number} 年度
 */
function getFiscalYearFromMonthKey(monthKey) {
  const parts = String(monthKey).split('-');
  if (parts.length !== 2) {
    throw new Error('monthKey の形式が不正です: ' + monthKey);
  }

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  if (!Number.isInteger(year) || !Number.isInteger(month)) {
    throw new Error('monthKey の値が不正です: ' + monthKey);
  }

  const date = new Date(year, month - 1, 1);
  return getFiscalYear(date);
}

/**
 * 月キー比較（yyyy-MM）
 * @param {string} a - 月キーA
 * @param {string} b - 月キーB
 * @return {number} 比較結果
 */
function monthKeyCompare(a, b) {
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

/**
 * 期間内の月キー一覧を生成
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array<string>} 月キー配列
 */
function listMonthKeysInRange(startDate, endDate) {
  const result = [];
  let cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const lastMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  while (cursor <= lastMonth) {
    result.push(formatMonthKey(cursor));
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }

  return result;
}

/**
 * 空の計画マップを作成
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 計画マップ
 */
function createEmptyPlanMap(startDate, endDate) {
  const map = { byMonth: {} };
  const monthKeys = listMonthKeysInRange(startDate, endDate);

  monthKeys.forEach(function(monthKey) {
    map.byMonth[monthKey] = {};
    for (let grade = 1; grade <= 6; grade++) {
      map.byMonth[monthKey][grade] = {
        planned_units: 0,
        school_days_count: 0,
        delta_units: 0,
        actual_units: 0,
        diff_units: 0
      };
    }
  });

  return map;
}

/**
 * 値が空欄かどうかを判定
 * @param {*} value - 判定対象
 * @return {boolean} 空欄でない場合true
 */
function isNonEmptyCell(value) {
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value === 'string') {
    return value.trim() !== '';
  }
  return value !== '';
}

/**
 * 日付を yyyy-MM 形式に変換
 * @param {Date} date - 対象日
 * @return {string} 月キー
 */
function formatMonthKey(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

/**
 * input[type=date] 用に yyyy-MM-dd 形式へ変換
 * @param {Date} date - 対象日
 * @return {string} 日付文字列
 */
function formatInputDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * 値を Date(00:00:00) へ正規化
 * @param {Date|string|number} value - 入力値
 * @return {Date|null} 正規化後の日付
 */
function normalizeToDate(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  let date = null;
  if (value instanceof Date) {
    date = new Date(value.getTime());
  } else if (typeof value === 'string') {
    const trimmed = value.trim();
    const ymd = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ymd) {
      date = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
    } else {
      date = new Date(trimmed);
    }
  } else {
    date = new Date(value);
  }

  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return null;
  }

  date.setHours(0, 0, 0, 0);
  return date;
}
