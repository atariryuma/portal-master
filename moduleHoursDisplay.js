/**
 * @fileoverview モジュール学習管理 - フォーマット・ユーティリティ・累計出力
 * @description 累計時数シートへの出力、表示フォーマット、日付/数値ユーティリティを担当します。
 */

/**
 * モジュール学習計画を集計し、累計時数シートへ統合出力
 * @param {Date|string} baseDate - 集計基準日
 * @param {?Object} options - 実行オプション（内部用）
 * @return {Object} 集計結果
 */
function syncModuleHoursWithCumulative(baseDate, options) {
  return syncModuleHoursWithCumulativeInternal(baseDate, options || null);
}

/**
 * モジュール学習計画を集計し、累計時数シートへ統合出力（内部）
 * @param {Date|string} baseDate - 集計基準日
 * @param {?Object} options - 実行オプション
 * @return {Object} 集計結果
 */
function syncModuleHoursWithCumulativeInternal(baseDate, options) {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const normalizedBaseDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const fiscalYear = getFiscalYear(normalizedBaseDate);
  const preservePlanningRange = options && options.preservePlanningRange ? options.preservePlanningRange : null;

  ensureDefaultCyclePlanForFiscalYear(fiscalYear, sheets.controlSheet);

  const buildResult = buildDailyPlanFromCyclePlanInternal(fiscalYear, normalizedBaseDate, {
    controlSheet: sheets.controlSheet
  });
  const exceptionTotals = loadExceptionTotals(fiscalYear, normalizedBaseDate, sheets.controlSheet);
  const gradeTotals = buildGradeTotalsFromDailyAndExceptions(buildResult.totalsByGrade, exceptionTotals);

  writeModuleToCumulativeSheet(gradeTotals, normalizedBaseDate);
  writeModulePlanSheet(buildResult, normalizedBaseDate);

  const settingsUpdates = {
    LAST_GENERATED_AT: new Date(),
    LAST_DAILY_PLAN_COUNT: buildResult.dailyPlanCount
  };
  if (preservePlanningRange && preservePlanningRange.startDate && preservePlanningRange.endDate) {
    settingsUpdates.PLAN_START_DATE = preservePlanningRange.startDate;
    settingsUpdates.PLAN_END_DATE = preservePlanningRange.endDate;
  } else {
    settingsUpdates.PLAN_START_DATE = buildResult.startDate;
    settingsUpdates.PLAN_END_DATE = buildResult.endDate;
  }
  upsertModuleSettingsValues(settingsUpdates);

  Logger.log('[INFO] モジュール学習計画を累計時数へ統合しました（基準日: ' + formatInputDate(normalizedBaseDate) + '）');

  return {
    baseDate: normalizedBaseDate,
    fiscalYear: fiscalYear,
    startDate: buildResult.startDate,
    endDate: buildResult.endDate,
    dailyPlanCount: buildResult.dailyPlanCount
  };
}

/**
 * 累計時数シートへモジュール累計を出力
 * @param {Object} gradeTotals - 学年別合計
 * @param {Date} baseDate - 基準日
 */
function writeModuleToCumulativeSheet(gradeTotals, baseDate) {
  const cumulativeSheet = getSheetByNameOrThrow(CUMULATIVE_SHEET.NAME);
  const displayColumn = MODULE_CUMULATIVE_COLUMNS.DISPLAY;
  const rowCount = MODULE_GRADE_MAX - MODULE_GRADE_MIN + 1;

  breakMergesInRange(cumulativeSheet, 2, MODULE_CUMULATIVE_COLUMNS.PLAN, rowCount + 1, displayColumn - MODULE_CUMULATIVE_COLUMNS.PLAN + 1);
  cleanupStaleDisplayColumns(cumulativeSheet, displayColumn, rowCount);

  cumulativeSheet
    .getRange(2, MODULE_CUMULATIVE_COLUMNS.PLAN, 1, 3)
    .setValues([['MOD計画累計', 'MOD実施累計', 'MOD差分']]);

  const valueRows = [];
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const total = gradeTotals[grade];
    valueRows.push([
      sessionsToUnits(total.elapsedPlannedSessions),
      sessionsToUnits(total.actualSessions),
      sessionsToUnits(total.diffSessions)
    ]);
  }

  cumulativeSheet.getRange(3, MODULE_CUMULATIVE_COLUMNS.PLAN, valueRows.length, 3).setValues(valueRows);

  cumulativeSheet.getRange(2, displayColumn).setValue(MODULE_DISPLAY_HEADER);
  const displayRows = [];
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    displayRows.push([buildModuleDisplayValue(gradeTotals[grade])]);
  }
  cumulativeSheet.getRange(3, displayColumn, displayRows.length, 1).setValues(displayRows);

  try {
    cumulativeSheet.hideColumns(MODULE_CUMULATIVE_COLUMNS.PLAN, 3);
  } catch (error) {
    Logger.log('[WARNING] MOD内部列の非表示に失敗: ' + error.toString());
  }
  try {
    cumulativeSheet.showColumns(displayColumn, 1);
  } catch (error) {
    Logger.log('[WARNING] MOD表示列の表示に失敗: ' + error.toString());
  }

  Logger.log('[INFO] モジュール表示列を更新しました（列: ' + displayColumn + ', 基準日: ' + formatInputDate(baseDate) + '）');
}

/**
 * モジュール計画シートへ日次配分スケジュールを出力
 * dailyRows を日付単位に集約し、カレンダー形式で書き込む。
 * @param {Object} buildResult - buildDailyPlanFromCyclePlanInternal の結果
 * @param {Date} baseDate - 集計基準日
 */
function writeModulePlanSheet(buildResult, baseDate) {
  const WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MODULE_SHEET_NAMES.SCHEDULE);
  if (!sheet) {
    sheet = ss.insertSheet(MODULE_SHEET_NAMES.SCHEDULE);
  }

  const cutoffDate = normalizeToDate(baseDate);
  const weekStart = getWeekStartMonday(cutoffDate);

  const dateMap = {};
  buildResult.dailyRows.forEach(function(row) {
    const dateKey = formatInputDate(row[0]);
    if (!dateMap[dateKey]) {
      dateMap[dateKey] = {
        date: row[0],
        cycleLabel: row[3],
        grades: {},
        elapsedFlag: row[7]
      };
    }
    const grade = Number(row[5]);
    dateMap[dateKey].grades[grade] = (dateMap[dateKey].grades[grade] || 0) + toNumberOrZero(row[6]);
  });

  const sortedKeys = Object.keys(dateMap).sort();
  const HEADER = ['日付', '曜日', 'クール', '1年', '2年', '3年', '4年', '5年', '6年', '状態'];
  const COL_COUNT = HEADER.length;

  const dataRows = sortedKeys.map(function(dateKey) {
    const entry = dateMap[dateKey];
    const date = entry.date;
    let status = '';
    if (entry.elapsedFlag === 1) {
      status = (date >= weekStart && date <= cutoffDate) ? '今週' : '済';
    }
    return [
      date,
      WEEKDAY_LABELS[date.getDay()],
      entry.cycleLabel,
      entry.grades[1] || '',
      entry.grades[2] || '',
      entry.grades[3] || '',
      entry.grades[4] || '',
      entry.grades[5] || '',
      entry.grades[6] || '',
      status
    ];
  });

  const allRows = [HEADER].concat(dataRows);
  const totalRows = allRows.length;

  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, totalRows, COL_COUNT).setValues(allRows);

  sheet.getRange(1, 1, 1, COL_COUNT)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 1).setNumberFormat('M/d');
    sheet.getRange(2, 2, dataRows.length, COL_COUNT - 1).setHorizontalAlignment('center');
  }

  Logger.log('[INFO] モジュール計画シートを更新しました（' + dataRows.length + '日分）');
}

/**
 * 指定範囲のセル結合を解除
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} startRow - 開始行
 * @param {number} startCol - 開始列
 * @param {number} numRows - 行数
 * @param {number} numCols - 列数
 */
function breakMergesInRange(sheet, startRow, startCol, numRows, numCols) {
  try {
    const range = sheet.getRange(startRow, startCol, numRows, numCols);
    const mergedRanges = range.getMergedRanges();
    for (let i = 0; i < mergedRanges.length; i++) {
      mergedRanges[i].breakApart();
    }
  } catch (error) {
    Logger.log('[WARNING] セル結合の解除に失敗: ' + error.toString());
  }
}

/**
 * 旧動的解決で作られた表示列の残骸をクリア（P列より右側）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cumulativeSheet - 累計時数
 * @param {number} displayColumn - 現在の表示列
 * @param {number} rowCount - データ行数
 */
function cleanupStaleDisplayColumns(cumulativeSheet, displayColumn, rowCount) {
  const lastColumn = cumulativeSheet.getLastColumn();
  if (lastColumn <= displayColumn) {
    return;
  }

  const extraColCount = lastColumn - displayColumn;
  const headers = cumulativeSheet.getRange(2, displayColumn + 1, 1, extraColCount).getValues()[0];
  const allExtraData = cumulativeSheet.getRange(3, displayColumn + 1, rowCount, extraColCount).getValues();

  for (let col = displayColumn + 1; col <= lastColumn; col++) {
    const colIndex = col - displayColumn - 1;
    const header = String(headers[colIndex] || '').trim();
    if (header === MODULE_DISPLAY_HEADER || header === '') {
      let hasDisplayData = false;
      for (let r = 0; r < allExtraData.length; r++) {
        const cellValue = String(allExtraData[r][colIndex] || '').trim();
        if (cellValue !== '' && cellValue.indexOf(MODULE_WEEKLY_LABEL) !== -1) {
          hasDisplayData = true;
          break;
        }
      }
      if (hasDisplayData) {
        cumulativeSheet.getRange(2, col, rowCount + 1, 1).clearContent();
        Logger.log('[INFO] 旧MOD表示列をクリアしました（列: ' + col + '）');
      }
    }
  }
}

/**
 * 表示列セル文字列を組み立て
 * @param {Object} total - 学年別合計
 * @return {string} 表示文字列
 */
function buildModuleDisplayValue(total) {
  return formatSessionsAsMixedFraction(total.actualSessions) +
    '（' + MODULE_WEEKLY_LABEL + ' ' + formatSignedSessionsAsMixedFraction(total.thisWeekSessions) + '）';
}

/**
 * 15分セッション数を整数+分数で表示
 * @param {number} sessions - セッション数
 * @return {string} 例: 18 2/3
 */
function formatSessionsAsMixedFraction(sessions) {
  const rounded = Math.round(toNumberOrZero(sessions));

  if (rounded === 0) {
    return '0';
  }

  const sign = rounded < 0 ? '-' : '';
  const absValue = Math.abs(rounded);
  const whole = Math.floor(absValue / 3);
  const remainder = absValue % 3;

  if (remainder === 0) {
    return sign + String(whole);
  }
  if (whole === 0) {
    return sign + remainder + '/3';
  }

  return sign + whole + ' ' + remainder + '/3';
}

/**
 * 符号付きの分数表示
 * @param {number} sessions - セッション数
 * @return {string} 例: +1/3
 */
function formatSignedSessionsAsMixedFraction(sessions) {
  const rounded = Math.round(toNumberOrZero(sessions));
  if (rounded > 0) {
    return '+' + formatSessionsAsMixedFraction(rounded);
  }
  return formatSessionsAsMixedFraction(rounded);
}

/**
 * セッション数を45分コマ数（小数）へ変換
 *
 * モジュール学習は15分単位（=1セッション）で記録される。
 * 累計時数シートは45分（=1コマ）単位のため、3セッション = 1コマで換算する。
 * 例: 21セッション = 7コマ、1セッション = 1/3コマ（≒0.333333）
 *
 * @param {number} sessions - セッション数（15分単位）
 * @return {number} 45分換算値
 */
function sessionsToUnits(sessions) {
  const value = toNumberOrZero(sessions) / 3;
  return Math.round(value * 1000000) / 1000000;
}

/**
 * 学年別合計の初期テンプレートを作成
 * @return {Object} 学年別合計
 */
function createGradeTotalsTemplate() {
  const result = {};

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    result[grade] = {
      plannedSessions: 0,
      elapsedPlannedSessions: 0,
      deltaSessions: 0,
      actualSessions: 0,
      diffSessions: 0,
      thisWeekSessions: 0
    };
  }

  return result;
}

/**
 * 年度の開始日・終了日を取得
 * @param {number} fiscalYear - 年度
 * @return {Object} 期間
 */
function getFiscalYearDateRange(fiscalYear) {
  const startDate = new Date(fiscalYear, MODULE_FISCAL_YEAR_START_MONTH - 1, 1);
  const endDate = new Date(fiscalYear + 1, MODULE_FISCAL_YEAR_START_MONTH - 1, 0);

  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  return {
    startDate: startDate,
    endDate: endDate
  };
}

/**
 * 期間が跨る年度一覧を取得
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array<number>} 年度一覧
 */
function collectFiscalYearsInRange(startDate, endDate) {
  const years = {};
  const cursor = new Date(startDate.getTime());

  while (cursor <= endDate) {
    years[getFiscalYear(cursor)] = true;
    cursor.setMonth(cursor.getMonth() + 1, 1);
  }

  return Object.keys(years).map(function(key) {
    return Number(key);
  }).sort();
}

/**
 * 保存済み期間を取得（未設定時は当該年度）
 * @param {Date} fallbackDate - 基準日
 * @param {Object=} settingsMap - 事前取得済み設定マップ
 * @return {Object} 期間
 */
function getModulePlanningRangeFromSettings(fallbackDate, settingsMap) {
  const map = settingsMap || readModuleSettingsMap();
  const start = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_START_DATE]);
  const end = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_END_DATE]);

  if (start && end && start <= end) {
    return { startDate: start, endDate: end };
  }

  const defaultRange = getDefaultModulePlanningRange(fallbackDate);
  upsertModuleSettingsValues({
    PLAN_START_DATE: defaultRange.startDate,
    PLAN_END_DATE: defaultRange.endDate
  });
  if (settingsMap) {
    settingsMap[MODULE_SETTING_KEYS.PLAN_START_DATE] = formatInputDate(defaultRange.startDate);
    settingsMap[MODULE_SETTING_KEYS.PLAN_END_DATE] = formatInputDate(defaultRange.endDate);
  }

  return defaultRange;
}

/**
 * 指定日を含む年度のデフォルト期間（4/1〜3/31）
 * @param {Date|string} baseDate - 基準日
 * @return {Object} 期間
 */
function getDefaultModulePlanningRange(baseDate) {
  const date = normalizeToDate(baseDate) || normalizeToDate(new Date());
  return getFiscalYearDateRange(getFiscalYear(date));
}

/**
 * 年度（4月開始）を取得
 *
 * 日本の学校年度は4月開始・3月終了。
 * 例: 2025年4月〜2026年3月は「2025年度」。
 * 1〜3月の日付は前年が年度値となる（2026年1月 → 2025年度）。
 *
 * @param {Date|string} date - 対象日
 * @param {number} startMonth - 年度開始月（デフォルト4月）
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
 * ダイアログ表示用に日時をフォーマット
 * @param {Date|string|number} value - 日時
 * @return {string} yyyy/MM/dd HH:mm 形式。無効時は未生成
 */
function formatDateTimeForDisplay(value) {
  if (value === null || value === undefined || value === '') {
    return '未生成';
  }

  const date = normalizeToDate(value);
  if (!date) {
    return '未生成';
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
}

/**
 * 値を Date(00:00:00) へ正規化
 * @param {Date|string|number} value - 入力値
 *   - Date: そのままコピーして時刻を 00:00:00 にリセット
 *   - string "yyyy-MM-dd": 年月日をパースして Date を生成
 *   - string (その他): new Date(value) にフォールバック（ランタイム依存）
 *   - number: Epoch ミリ秒として new Date(value) で変換
 *   - null/undefined/空文字: null を返す
 * @return {Date|null} 正規化後の日付。パース不能時は null
 */
function normalizeToDate(value) {
  const date = parseDateValue_(value);
  if (!date) {
    return null;
  }
  date.setHours(0, 0, 0, 0);
  return date;
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
 * 数値へ安全変換
 * @param {*} value - 入力値
 * @return {number} 数値（変換不能時0）
 */
function toNumberOrZero(value) {
  const numeric = Number(value);
  return isNaN(numeric) ? 0 : numeric;
}
