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

  const buildResult = buildDailyPlanFromCyclePlanInternal(fiscalYear, normalizedBaseDate, false, {
    controlSheet: sheets.controlSheet
  });
  const exceptionTotals = loadExceptionTotals(fiscalYear, normalizedBaseDate, sheets.controlSheet);
  const gradeTotals = buildGradeTotalsFromDailyAndExceptions(buildResult.totalsByGrade, exceptionTotals);

  writeModuleSummary(gradeTotals, fiscalYear, normalizedBaseDate, null);
  writeModuleToCumulativeSheet(gradeTotals, normalizedBaseDate, null);

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
  upsertModuleSettingsValues(null, settingsUpdates);

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
 * 旧互換: summary 出力（現在は保存せずログのみ）
 * @param {Object} gradeTotals - 学年別合計
 * @param {number} fiscalYear - 対象年度
 * @param {Date} baseDate - 基準日
 */
function writeModuleSummary(gradeTotals, fiscalYear, baseDate) {
  Logger.log('[INFO] module_summary への保存は廃止済み（FY' + fiscalYear + ', 基準日: ' + formatInputDate(baseDate) + '）');
  return gradeTotals;
}

/**
 * 累計時数シートへモジュール累計を出力
 * @param {Object} gradeTotals - 学年別合計
 * @param {Date} baseDate - 基準日
 */
function writeModuleToCumulativeSheet(gradeTotals, baseDate) {
  const cumulativeSheet = getSheetByNameOrThrow(CUMULATIVE_SHEET.NAME);

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

  const displayColumn = resolveCumulativeDisplayColumn(cumulativeSheet);
  const displayRows = [];
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    displayRows.push([buildModuleDisplayValue(gradeTotals[grade])]);
  }
  cumulativeSheet.getRange(3, displayColumn, displayRows.length, 1).setValues(displayRows);
  enforceModuleCumulativeColumnVisibility(cumulativeSheet, displayColumn);

  upsertModuleSettingsValues(null, {
    CUMULATIVE_DISPLAY_COLUMN: displayColumn
  });

  Logger.log('[INFO] モジュール表示列を更新しました（列: ' + displayColumn + ', 基準日: ' + formatInputDate(baseDate) + '）');
}

/**
 * モジュール累計の表示列運用を整える（M〜Oは非表示、表示列は可視）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cumulativeSheet - 累計時数
 * @param {number} displayColumn - 表示列
 */
function enforceModuleCumulativeColumnVisibility(cumulativeSheet, displayColumn) {
  try {
    cumulativeSheet.hideColumns(MODULE_CUMULATIVE_COLUMNS.PLAN, 3);
  } catch (error) {
    Logger.log('[WARNING] MOD内部列の非表示に失敗: ' + error.toString());
  }

  const column = Number(displayColumn);
  if (!Number.isInteger(column) || column < 1) {
    return;
  }

  try {
    cumulativeSheet.showColumns(column, 1);
  } catch (error) {
    Logger.log('[WARNING] MOD表示列の表示に失敗: ' + error.toString());
  }
}

/**
 * 累計時数の表示列を動的に解決
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cumulativeSheet - 累計時数
 * @param {Object=} settingsMap - 事前取得済み設定マップ
 * @return {number} 列番号（1-based）
 */
function resolveCumulativeDisplayColumn(cumulativeSheet, settingsMap) {
  const map = settingsMap || readModuleSettingsMap();
  const displayRowCount = MODULE_GRADE_MAX - MODULE_GRADE_MIN + 1;
  const preferredColumn = MODULE_CUMULATIVE_COLUMNS.DISPLAY_FALLBACK;

  const preferredHeader = String(cumulativeSheet.getRange(2, preferredColumn).getValue() || '').trim();
  if (preferredHeader === MODULE_DISPLAY_HEADER ||
      (preferredHeader === '' &&
        isReusableCumulativeDisplayColumn(cumulativeSheet, preferredColumn, displayRowCount))) {
    cumulativeSheet.getRange(2, preferredColumn).setValue(MODULE_DISPLAY_HEADER);
    return preferredColumn;
  }

  const configuredColumn = Number(map[MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN]);
  if (Number.isInteger(configuredColumn) && configuredColumn >= 1) {
    const configuredHeader = String(cumulativeSheet.getRange(2, configuredColumn).getValue() || '').trim();
    if ((configuredHeader === '' || configuredHeader === MODULE_DISPLAY_HEADER) &&
        isReusableCumulativeDisplayColumn(cumulativeSheet, configuredColumn, displayRowCount)) {
      cumulativeSheet.getRange(2, configuredColumn).setValue(MODULE_DISPLAY_HEADER);
      return configuredColumn;
    }
  }

  const fallbackStart = MODULE_CUMULATIVE_COLUMNS.DISPLAY_FALLBACK;
  const lastColumn = Math.max(cumulativeSheet.getLastColumn(), fallbackStart);
  let emptyColumn = null;

  for (let col = fallbackStart; col <= lastColumn; col++) {
    const header = String(cumulativeSheet.getRange(2, col).getValue() || '').trim();
    if (header === MODULE_DISPLAY_HEADER) {
      return col;
    }
    if (!emptyColumn && header === '' &&
        isReusableCumulativeDisplayColumn(cumulativeSheet, col, displayRowCount)) {
      emptyColumn = col;
    }
  }

  const resolved = emptyColumn || (lastColumn + 1);
  cumulativeSheet.getRange(2, resolved).setValue(MODULE_DISPLAY_HEADER);
  return resolved;
}

/**
 * 累計表示列が再利用可能か判定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} cumulativeSheet - 累計時数
 * @param {number} column - 対象列（1-based）
 * @param {number} rowCount - 確認行数
 * @return {boolean} 再利用可能なら true
 */
function isReusableCumulativeDisplayColumn(cumulativeSheet, column, rowCount) {
  if (rowCount <= 0) {
    return true;
  }

  const values = cumulativeSheet.getRange(3, column, rowCount, 1).getValues();
  return values.every(function(row) {
    return !isNonEmptyCell(row[0]);
  });
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
 * @param {number} sessions - セッション数
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
 * 学年別セッション合算
 * @param {Object} gradeTotals - 学年別合計
 * @param {number} grade - 学年
 * @param {number} sessions - 加算値
 * @param {string} field - 項目名
 */
function addGradeSessions(gradeTotals, grade, sessions, field) {
  if (!gradeTotals[grade]) {
    return;
  }
  gradeTotals[grade][field] = toNumberOrZero(gradeTotals[grade][field]) + toNumberOrZero(sessions);
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
 * @param {*} settingsSheet - 旧互換引数（未使用）
 * @param {Date} fallbackDate - 基準日
 * @param {Object=} settingsMap - 事前取得済み設定マップ
 * @return {Object} 期間
 */
function getModulePlanningRangeFromSettings(settingsSheet, fallbackDate, settingsMap) {
  const map = settingsMap || readModuleSettingsMap();
  const start = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_START_DATE]);
  const end = normalizeToDate(map[MODULE_SETTING_KEYS.PLAN_END_DATE]);

  if (start && end && start <= end) {
    return { startDate: start, endDate: end };
  }

  const defaultRange = getDefaultModulePlanningRange(fallbackDate);
  upsertModuleSettingsValues(null, {
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

  return getFiscalYear(new Date(year, month - 1, 1));
}

/**
 * 月キー比較（yyyy-MM）
 * @param {string} a - 月キーA
 * @param {string} b - 月キーB
 * @return {number} 比較結果
 */
function monthKeyCompare(a, b) {
  if (a === b) {
    return 0;
  }
  return a < b ? -1 : 1;
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

  const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '未生成';
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
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
