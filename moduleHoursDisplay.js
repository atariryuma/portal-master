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
  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  const normalizedBaseDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const fiscalYear = getFiscalYear(normalizedBaseDate);
  const preservePlanningRange = options && options.preservePlanningRange ? options.preservePlanningRange : null;

  ensureDefaultAnnualTargetForFiscalYear(fiscalYear, controlSheet);

  const settingsMap = readModuleSettingsMap();
  const enabledWeekdays = getEnabledWeekdays(settingsMap);
  const planningRange = preservePlanningRange
    ? { startDate: normalizeToDate(preservePlanningRange.startDate), endDate: normalizeToDate(preservePlanningRange.endDate) }
    : getModulePlanningRangeFromSettings(normalizedBaseDate, settingsMap);

  const buildResult = buildDailyPlanFromAnnualTarget(fiscalYear, normalizedBaseDate, {
    controlSheet: controlSheet,
    enabledWeekdays: enabledWeekdays,
    startDate: planningRange.startDate,
    endDate: planningRange.endDate
  });
  const exceptionTotals = loadExceptionTotals(fiscalYear, normalizedBaseDate, controlSheet);
  const gradeTotals = buildGradeTotalsFromDailyAndExceptions(buildResult.totalsByGrade, exceptionTotals);

  writeModuleToCumulativeSheet(gradeTotals, normalizedBaseDate);

  const annualTarget = loadAnnualTargetForFiscalYear(fiscalYear, controlSheet);
  writeModulePlanSummarySheet(buildResult, annualTarget, enabledWeekdays, normalizedBaseDate);

  upsertModuleSettingsValues({
    LAST_GENERATED_AT: new Date(),
    LAST_DAILY_PLAN_COUNT: buildResult.dailyPlanCount,
    PLAN_START_DATE: planningRange.startDate,
    PLAN_END_DATE: planningRange.endDate
  });

  Logger.log('[INFO] モジュール学習計画を累計時数へ統合しました（基準日: ' + formatInputDate(normalizedBaseDate) + '）');

  return {
    baseDate: normalizedBaseDate,
    fiscalYear: fiscalYear,
    startDate: planningRange.startDate,
    endDate: planningRange.endDate,
    dailyPlanCount: buildResult.dailyPlanCount,
    reserveByGrade: buildResult.reserveByGrade
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
 * モジュール学習計画シートへ年間実施計画を出力
 * @param {Object} buildResult - buildDailyPlanFromAnnualTarget の返却値
 * @param {Object} annualTarget - 年間目標 { gradeKoma }
 * @param {Array<number>} enabledWeekdays - 有効曜日配列
 * @param {Date} baseDate - 基準日
 */
function writeModulePlanSummarySheet(buildResult, annualTarget, enabledWeekdays, baseDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreatePlanSummarySheet(ss);
    if (!sheet) {
      return;
    }

    const fiscalYear = buildResult.fiscalYear;
    const startMonth = MODULE_FISCAL_YEAR_START_MONTH;

    // 月別×学年別セッション集計
    const monthlyByGrade = {};
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      monthlyByGrade[grade] = {};
    }
    buildResult.dailyRows.forEach(function(row) {
      const date = normalizeToDate(row[0]);
      const grade = Number(row[5]);
      const sessions = toNumberOrZero(row[6]);
      if (!date || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
        return;
      }
      const monthKey = formatMonthKey(date);
      monthlyByGrade[grade][monthKey] = toNumberOrZero(monthlyByGrade[grade][monthKey]) + sessions;
    });

    // 年度月順（4月→3月）のキーリスト
    const monthKeys = [];
    for (let i = 0; i < 12; i++) {
      const m = ((startMonth - 1 + i) % 12) + 1;
      const y = m >= startMonth ? fiscalYear : fiscalYear + 1;
      monthKeys.push(String(y) + '-' + String(m).padStart(2, '0'));
    }
    const monthLabels = monthKeys.map(function(key) {
      return String(parseInt(key.split('-')[1], 10)) + '月';
    });

    // 曜日ラベル
    const weekdayNames = enabledWeekdays
      .slice().sort(function(a, b) { return a - b; })
      .map(function(d) { return MODULE_WEEKDAY_LABELS[d] || String(d); })
      .join('・');

    // シートクリア・書式リセット
    sheet.clear();
    sheet.clearFormats();

    // 行1: タイトル
    const titleRange = sheet.getRange(1, 1, 1, 16);
    titleRange.merge();
    sheet.getRange(1, 1).setValue('モジュール学習 年間実施計画');
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

    // 行3: 年度・実施期間
    sheet.getRange(3, 1).setValue(
      '年度: ' + fiscalYear + '年度　　実施期間: ' +
      formatInputDate(buildResult.startDate) + ' ～ ' + formatInputDate(buildResult.endDate)
    );

    // 行4: 実施曜日・形式
    sheet.getRange(4, 1).setValue(
      '実施曜日: ' + weekdayNames + '　　1回15分（3回で1単位時間 = 45分）'
    );

    // 行6: ヘッダー
    const headerRow = 6;
    const headers = ['学年'];
    monthLabels.forEach(function(label) { headers.push(label); });
    headers.push('合計');
    headers.push('年間目標');
    headers.push('予備/不足');

    sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(headerRow, 1, 1, headers.length)
      .setBackground('#f0f0f0')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true);

    // 行7-12: 学年データ
    const dataStartRow = headerRow + 1;
    const dataRows = [];
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const row = [grade + '年'];
      let totalSessions = 0;

      monthKeys.forEach(function(monthKey) {
        const sessions = toNumberOrZero(monthlyByGrade[grade][monthKey]);
        totalSessions += sessions;
        row.push(sessions > 0 ? formatSessionsAsMixedFraction(sessions) : '');
      });

      row.push(formatSessionsAsMixedFraction(totalSessions));
      row.push(toNumberOrZero(annualTarget.gradeKoma[grade]));

      const reserve = toNumberOrZero(buildResult.reserveByGrade[grade]);
      if (reserve > 0) {
        row.push(MODULE_RESERVE_LABEL + ' ' + formatSessionsAsMixedFraction(reserve) + 'コマ');
      } else if (reserve < 0) {
        row.push(MODULE_DEFICIT_LABEL + ' ' + formatSessionsAsMixedFraction(Math.abs(reserve)) + 'コマ');
      } else {
        row.push('-');
      }

      dataRows.push(row);
    }

    sheet.getRange(dataStartRow, 1, dataRows.length, headers.length).setValues(dataRows);

    // データ行の書式
    const dataRange = sheet.getRange(dataStartRow, 1, dataRows.length, headers.length);
    dataRange.setBorder(true, true, true, true, true, true);
    sheet.getRange(dataStartRow, 2, dataRows.length, 12).setHorizontalAlignment('center');
    sheet.getRange(dataStartRow, 14, dataRows.length, 3).setHorizontalAlignment('center');

    // 不足セルに赤背景
    const reserveCol = headers.length;
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const reserveVal = toNumberOrZero(buildResult.reserveByGrade[grade]);
      const rowIndex = dataStartRow + grade - MODULE_GRADE_MIN;
      if (reserveVal < 0) {
        sheet.getRange(rowIndex, reserveCol).setBackground('#fef2f2').setFontColor('#991b1b').setFontWeight('bold');
      }
    }

    // ── 日別実施計画セクション ──
    const dailySectionStartRow = dataStartRow + dataRows.length + 2;
    let lastContentRow = dataStartRow + dataRows.length;

    const dailyByDate = {};
    buildResult.dailyRows.forEach(function(row) {
      const date = normalizeToDate(row[0]);
      if (!date) {
        return;
      }
      const dateKey = formatInputDate(date);
      if (!dailyByDate[dateKey]) {
        dailyByDate[dateKey] = { date: date, grades: {} };
      }
      dailyByDate[dateKey].grades[Number(row[5])] = toNumberOrZero(row[6]);
    });

    const sortedDateKeys = Object.keys(dailyByDate).sort();
    if (sortedDateKeys.length > 0) {
      sheet.getRange(dailySectionStartRow, 1).setValue('日別実施計画');
      sheet.getRange(dailySectionStartRow, 1).setFontSize(12).setFontWeight('bold');

      const dailyHeaderRow = dailySectionStartRow + 1;
      const dailyHeaders = ['日付', '曜日'];
      for (let g = MODULE_GRADE_MIN; g <= MODULE_GRADE_MAX; g++) {
        dailyHeaders.push(g + '年');
      }
      sheet.getRange(dailyHeaderRow, 1, 1, dailyHeaders.length).setValues([dailyHeaders]);
      sheet.getRange(dailyHeaderRow, 1, 1, dailyHeaders.length)
        .setBackground('#f0f0f0')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true);

      const dailyDataRows = [];
      const monthBoundaryIndices = [];
      let prevMonth = -1;
      sortedDateKeys.forEach(function(dateKey) {
        const entry = dailyByDate[dateKey];
        const date = entry.date;
        const month = date.getMonth() + 1;
        if (prevMonth !== -1 && month !== prevMonth) {
          monthBoundaryIndices.push(dailyDataRows.length);
        }
        prevMonth = month;

        const dayOfWeek = date.getDay();
        const weekdayLabel = MODULE_WEEKDAY_LABELS[dayOfWeek] || '';
        const dailyRow = [String(month) + '/' + String(date.getDate()), weekdayLabel];
        for (let g = MODULE_GRADE_MIN; g <= MODULE_GRADE_MAX; g++) {
          const sessions = entry.grades[g] || 0;
          dailyRow.push(sessions > 0 ? sessions : '');
        }
        dailyDataRows.push(dailyRow);
      });

      if (dailyDataRows.length > 0) {
        const dailyDataStartRow = dailyHeaderRow + 1;
        sheet.getRange(dailyDataStartRow, 1, dailyDataRows.length, dailyHeaders.length).setValues(dailyDataRows);
        sheet.getRange(dailyDataStartRow, 1, dailyDataRows.length, dailyHeaders.length)
          .setBorder(true, true, true, true, true, true);
        sheet.getRange(dailyDataStartRow, 2, dailyDataRows.length, dailyHeaders.length - 1)
          .setHorizontalAlignment('center');

        monthBoundaryIndices.forEach(function(rowIdx) {
          if (rowIdx > 0) {
            sheet.getRange(dailyDataStartRow + rowIdx - 1, 1, 1, dailyHeaders.length)
              .setBorder(null, null, true, null, null, null, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          }
        });

        lastContentRow = dailyDataStartRow + dailyDataRows.length - 1;
      }
    }

    // フッター行
    const footerRow = lastContentRow + 2;
    sheet.getRange(footerRow, 1).setValue(
      '更新日時: ' + formatDateTimeForDisplay(new Date()) + '　　※ 本シートは再集計時に自動更新されます'
    );
    sheet.getRange(footerRow, 1).setFontSize(9).setFontColor('#64748b');

    // 列幅調整
    sheet.setColumnWidth(1, 50);
    for (let c = 2; c <= 13; c++) {
      sheet.setColumnWidth(c, 45);
    }
    sheet.setColumnWidth(14, 55);
    sheet.setColumnWidth(15, 55);
    sheet.setColumnWidth(16, 75);

    Logger.log('[INFO] モジュール学習計画シートを更新しました');
  } catch (error) {
    Logger.log('[WARNING] モジュール学習計画シートの更新に失敗: ' + error.toString());
  }
}

/**
 * モジュール学習計画シートを取得または作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シート
 */
function getOrCreatePlanSummarySheet(ss) {
  const sheet = ss.getSheetByName(MODULE_SHEET_NAMES.PLAN_SUMMARY);
  if (sheet) {
    return sheet;
  }
  return ss.insertSheet(MODULE_SHEET_NAMES.PLAN_SUMMARY);
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

  const checkColCount = lastColumn - displayColumn;
  const headers = cumulativeSheet.getRange(2, displayColumn + 1, 1, checkColCount).getValues()[0];

  for (let i = 0; i < checkColCount; i++) {
    const header = String(headers[i] || '').trim();
    const col = displayColumn + 1 + i;
    if (header === MODULE_DISPLAY_HEADER || header === '') {
      let hasDisplayData = false;
      const values = cumulativeSheet.getRange(3, col, rowCount, 1).getValues();
      for (let r = 0; r < values.length; r++) {
        const cellValue = String(values[r][0] || '').trim();
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
      thisWeekSessions: 0,
      reserveSessions: 0
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
