/**
 * @fileoverview モジュール学習管理機能
 * @description module_control 1シートに計画入力と例外入力を集約し、
 *              日次配分は再計算時にメモリ上で生成して累計時数へ統合します。
 */

const MODULE_DEFAULT_CYCLES = [
  { order: 1, startMonth: 6, endMonth: 7, label: '6-7' },
  { order: 2, startMonth: 9, endMonth: 10, label: '9-10' },
  { order: 3, startMonth: 11, endMonth: 12, label: '11-12' },
  { order: 4, startMonth: 1, endMonth: 2, label: '1-2' }
];

const MODULE_DEFAULT_KOMA_PER_CYCLE = 7;
const MODULE_DISPLAY_HEADER = 'MOD実施累計(表示)';
const MODULE_WEEKLY_LABEL = '今週';
const MODULE_GRADE_MIN = 1;
const MODULE_GRADE_MAX = 6;
const MODULE_SETTINGS_PREFIX = 'MODULE_';

const MODULE_WEEKDAY_PRIORITY = {
  1: 0, // 月
  3: 1, // 水
  5: 2, // 金
  2: 3, // 火
  4: 4  // 木
};

const MODULE_CONTROL_MARKERS = {
  PLAN: 'PLAN_TABLE',
  EXCEPTIONS: 'EXCEPTIONS_TABLE'
};

const MODULE_CONTROL_DEFAULT_LAYOUT = {
  VERSION_ROW: 1,
  PLAN_MARKER_ROW: 3,
  EXCEPTIONS_MARKER_ROW: 40
};

const MODULE_CONTROL_PLAN_HEADERS = [
  'fiscal_year',
  'cycle_order',
  'start_month',
  'end_month',
  'g1_koma',
  'g2_koma',
  'g3_koma',
  'g4_koma',
  'g5_koma',
  'g6_koma',
  'note'
];

const MODULE_CONTROL_EXCEPTION_HEADERS = [
  'date',
  'grade',
  'delta_sessions',
  'reason',
  'note'
];

/**
 * モジュール学習管理ダイアログを表示
 */
function showModulePlanningDialog() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('modulePlanningDialog')
      .setWidth(980)
      .setHeight(720);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'モジュール学習管理');
  } catch (error) {
    showAlert('モジュール学習管理ダイアログの表示に失敗しました: ' + error.toString(), 'エラー');
  }
}

/**
 * ダイアログ用の旧互換初期値（期間）
 * @return {Object} 開始日・終了日
 */
function getModulePlanningDefaults() {
  const state = getModulePlanningDialogState();
  return {
    startDate: state.startDate,
    endDate: state.endDate
  };
}

/**
 * ダイアログ表示用の状態を返却
 * @return {Object} ダイアログ状態
 */
function getModulePlanningDialogState() {
  const startedAt = new Date().getTime();
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const baseDate = normalizeToDate(getCurrentOrNextSaturday());
  const fiscalYear = getFiscalYear(baseDate);
  const fiscalRange = getFiscalYearDateRange(fiscalYear);

  ensureDefaultCyclePlanForFiscalYear(fiscalYear, sheets.controlSheet);

  const settingsMap = readModuleSettingsMap();
  const savedRange = getModulePlanningRangeFromSettings(null, baseDate);
  const dailyPlanCount = getCachedDailyPlanCountForDialog(settingsMap);
  const cyclePlans = buildDialogCyclePlansForFiscalYear(fiscalYear, sheets.controlSheet);
  const recentExceptions = listRecentExceptionsForFiscalYear(sheets.controlSheet, fiscalYear, 10);

  try {
    const cumulativeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('累計時数');
    if (cumulativeSheet) {
      const displayColumn = resolveCumulativeDisplayColumn(cumulativeSheet);
      enforceModuleCumulativeColumnVisibility(cumulativeSheet, displayColumn);
    }
  } catch (error) {
    Logger.log('[WARNING] 累計時数の列表示制御に失敗: ' + error.toString());
  }

  const state = {
    baseDate: formatInputDate(baseDate),
    fiscalYear: fiscalYear,
    fiscalYearStartDate: formatInputDate(fiscalRange.startDate),
    fiscalYearEndDate: formatInputDate(fiscalRange.endDate),
    startDate: formatInputDate(savedRange.startDate),
    endDate: formatInputDate(savedRange.endDate),
    lastGeneratedAt: formatDateTimeForDisplay(settingsMap[MODULE_SETTING_KEYS.LAST_GENERATED_AT]),
    cyclePlanRecordCount: countCyclePlanRowsForFiscalYear(sheets.controlSheet, fiscalYear),
    dailyPlanRecordCount: dailyPlanCount,
    exceptionRecordCount: countExceptionRowsForFiscalYear(sheets.controlSheet, fiscalYear),
    cumulativeDisplayColumn: settingsMap[MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN] || '',
    cyclePlans: cyclePlans,
    recentExceptions: recentExceptions
  };

  const elapsedMs = new Date().getTime() - startedAt;
  if (elapsedMs >= 2000) {
    Logger.log('[PERF] getModulePlanningDialogState: ' + elapsedMs + 'ms');
  }

  return state;
}

/**
 * ダイアログ用の日次件数キャッシュ値を取得
 * @param {Object} settingsMap - 設定マップ
 * @return {number} 件数
 */
function getCachedDailyPlanCountForDialog(settingsMap) {
  const value = Number(settingsMap[MODULE_SETTING_KEYS.LAST_DAILY_PLAN_COUNT]);
  if (!Number.isFinite(value) || value < 0) {
    return 0;
  }
  return Math.round(value);
}

/**
 * ダイアログ表示用のクール計画を取得
 * @param {number} fiscalYear - 対象年度
 * @return {Array<Object>} 計画配列
 */
function buildDialogCyclePlansForFiscalYear(fiscalYear, controlSheet) {
  const plans = loadCyclePlanForFiscalYear(fiscalYear, controlSheet);
  return plans.map(function(plan) {
    return {
      cycleOrder: plan.cycleOrder,
      startMonth: plan.startMonth,
      endMonth: plan.endMonth,
      g1Koma: plan.gradeKoma[1],
      g2Koma: plan.gradeKoma[2],
      g3Koma: plan.gradeKoma[3],
      g4Koma: plan.gradeKoma[4],
      g5Koma: plan.gradeKoma[5],
      g6Koma: plan.gradeKoma[6],
      note: plan.note || ''
    };
  });
}

/**
 * 対象年度の最近の実施差分を返却
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @param {number} limitCount - 取得件数
 * @return {Array<Object>} 実施差分配列
 */
function listRecentExceptionsForFiscalYear(controlSheet, fiscalYear, limitCount) {
  const limit = Math.max(1, Number(limitCount) || 10);
  const targetFiscalYear = Number(fiscalYear);

  return readExceptionRows(controlSheet)
    .map(function(item) {
      return {
        rowNumber: item.rowNumber,
        date: normalizeToDate(item.date),
        grade: Number(item.grade),
        deltaSessions: Math.round(toNumberOrZero(item.deltaSessions)),
        reason: item.reason || '',
        note: item.note || ''
      };
    })
    .filter(function(item) {
      return !!item.date &&
        getFiscalYear(item.date) === targetFiscalYear &&
        Number.isInteger(item.grade) &&
        item.grade >= MODULE_GRADE_MIN &&
        item.grade <= MODULE_GRADE_MAX;
    })
    .sort(function(a, b) {
      if (a.date.getTime() !== b.date.getTime()) {
        return b.date.getTime() - a.date.getTime();
      }
      return b.rowNumber - a.rowNumber;
    })
    .slice(0, limit)
    .map(function(item) {
      return {
        date: formatInputDate(item.date),
        grade: item.grade,
        deltaSessions: item.deltaSessions,
        deltaDisplay: formatSignedSessionsAsMixedFraction(item.deltaSessions),
        reason: item.reason,
        note: item.note
      };
    });
}

/**
 * 旧互換: 管理画面を開く
 * @param {string=} section - 表示セクション（plan / exceptions）
 * @return {string} 完了メッセージ
 */
function openModuleControlSheet(section) {
  showModulePlanningDialog();
  if (section === 'exceptions') {
    return 'モジュール学習管理を開きました（実施差分入力）。';
  }
  return 'モジュール学習管理を開きました（計画入力）。';
}

/**
 * 旧互換: cycle 計画シートを開く
 * @return {string} 完了メッセージ
 */
function openModuleCyclePlanSheet() {
  return openModuleControlSheet('plan');
}

/**
 * 旧互換: daily 計画シートを開く
 * @return {string} 完了メッセージ
 */
function openModuleDailyPlanSheet() {
  return openModuleControlSheet('exceptions');
}

/**
 * モジュール学習集計を再実行
 * @return {string} 完了メッセージ
 */
function refreshModulePlanning() {
  const baseDate = getCurrentOrNextSaturday();
  const result = syncModuleHoursWithCumulative(baseDate);
  return [
    'モジュール学習集計を更新しました。',
    '基準日: ' + formatInputDate(result.baseDate),
    '対象年度: ' + result.fiscalYear + '年度',
    '日次計画件数（再計算）: ' + result.dailyPlanCount + '件'
  ].join('\n');
}

/**
 * ダイアログから受け取ったクール計画を保存して再集計
 * @param {Object} payload - 入力データ
 * @return {string} 完了メッセージ
 */
function saveModuleCyclePlanFromDialog(payload) {
  const fiscalYear = Number(payload && payload.fiscalYear);
  if (!Number.isInteger(fiscalYear) || fiscalYear < 2000 || fiscalYear > 2100) {
    throw new Error('対象年度が不正です。');
  }

  const plans = payload && Array.isArray(payload.plans) ? payload.plans : [];
  const rows = normalizeCyclePlanRowsFromDialog(fiscalYear, plans);
  if (rows.length === 0) {
    throw new Error('保存対象のクール計画がありません。');
  }

  const sheets = initializeModuleHoursSheetsIfNeeded();
  replaceCyclePlanRowsForFiscalYearInControl(sheets.controlSheet, fiscalYear, rows);

  const baseDate = normalizeToDate(payload && payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const result = syncModuleHoursWithCumulative(baseDate);

  return [
    '計画を保存して再集計しました。',
    '対象年度: ' + fiscalYear + '年度',
    '計画件数: ' + rows.length + '件',
    '基準日: ' + formatInputDate(result.baseDate)
  ].join('\n');
}

/**
 * ダイアログから実施差分を追加して再集計
 * @param {Object} payload - 入力データ
 * @return {string} 完了メッセージ
 */
function addModuleExceptionFromDialog(payload) {
  const exceptionDate = normalizeToDate(payload && payload.date);
  if (!exceptionDate) {
    throw new Error('日付が不正です。');
  }

  const grade = Number(payload && payload.grade);
  if (!Number.isInteger(grade) || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
    throw new Error('学年は1〜6で入力してください。');
  }

  const deltaSessions = Math.round(toNumberOrZero(payload && payload.deltaSessions));
  if (!Number.isFinite(deltaSessions) || deltaSessions === 0) {
    throw new Error('調整値は0以外の数値を入力してください。');
  }

  const reason = String(payload && payload.reason ? payload.reason : '').trim();
  const note = String(payload && payload.note ? payload.note : '').trim();

  const sheets = initializeModuleHoursSheetsIfNeeded();
  appendExceptionRows(sheets.controlSheet, [[exceptionDate, grade, deltaSessions, reason, note]]);

  const baseDate = normalizeToDate(payload && payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const result = syncModuleHoursWithCumulative(baseDate);

  return [
    '実施差分を保存して再集計しました。',
    '入力: ' + formatInputDate(exceptionDate) + ' / ' + grade + '年 / ' + formatSignedSessionsAsMixedFraction(deltaSessions),
    '基準日: ' + formatInputDate(result.baseDate)
  ].join('\n');
}

/**
 * ダイアログ入力値をクール計画行へ正規化
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Object>} plans - 入力計画
 * @return {Array<Array<*>>} シート行
 */
function normalizeCyclePlanRowsFromDialog(fiscalYear, plans) {
  const rows = [];
  const seenOrder = {};

  plans.forEach(function(plan, index) {
    const cycleOrder = Number(plan && plan.cycleOrder);
    const startMonth = Number(plan && plan.startMonth);
    const endMonth = Number(plan && plan.endMonth);

    if (!Number.isInteger(cycleOrder) || cycleOrder <= 0) {
      throw new Error('クール順が不正です（行 ' + (index + 1) + '）。');
    }
    if (seenOrder[cycleOrder]) {
      throw new Error('クール順が重複しています: ' + cycleOrder);
    }
    if (!isValidModuleMonth(startMonth) || !isValidModuleMonth(endMonth)) {
      throw new Error('開始月または終了月が不正です（クール ' + cycleOrder + '）。');
    }

    const gradeValues = [];
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const key = 'g' + grade + 'Koma';
      const rawValue = toNumberOrZero(plan && plan[key]);
      if (rawValue < 0) {
        throw new Error(grade + '年のコマ数は0以上で入力してください（クール ' + cycleOrder + '）。');
      }
      gradeValues.push(Math.round(rawValue * 1000) / 1000);
    }

    const note = String(plan && plan.note ? plan.note : '').trim();

    rows.push([
      Number(fiscalYear),
      cycleOrder,
      startMonth,
      endMonth,
      gradeValues[0],
      gradeValues[1],
      gradeValues[2],
      gradeValues[3],
      gradeValues[4],
      gradeValues[5],
      note
    ]);
    seenOrder[cycleOrder] = true;
  });

  rows.sort(function(a, b) {
    return Number(a[1]) - Number(b[1]);
  });

  return rows;
}

/**
 * 月値の妥当性を判定
 * @param {number} month - 月
 * @return {boolean} 妥当なら true
 */
function isValidModuleMonth(month) {
  return Number.isInteger(month) && month >= 1 && month <= 12;
}

/**
 * 旧期間入力APIの後方互換ラッパー
 * @param {string|Date} startDate - 開始日
 * @param {string|Date} endDate - 終了日
 * @return {string} 完了メッセージ
 */
function saveModulePlanningRange(startDate, endDate) {
  const result = rebuildModulePlanFromRange(startDate, endDate);
  return [
    'モジュール学習計画を更新しました。',
    '対象期間: ' + formatInputDate(result.startDate) + ' ～ ' + formatInputDate(result.endDate),
    '対象年度: ' + result.fiscalYears.join(', '),
    '生成件数: ' + result.recordCount + '件',
    '※ 設定はモジュール学習管理画面で一元管理します。'
  ].join('\n');
}

/**
 * 旧期間指定で計画再生成（後方互換）
 * @param {string|Date} startDate - 開始日
 * @param {string|Date} endDate - 終了日
 * @return {Object} 再生成結果
 */
function rebuildModulePlanFromRange(startDate, endDate) {
  const start = normalizeToDate(startDate);
  const end = normalizeToDate(endDate);

  if (!start || !end) {
    throw new Error('開始日・終了日は yyyy-MM-dd 形式で入力してください。');
  }
  if (start > end) {
    throw new Error('開始日は終了日以前の日付を指定してください。');
  }

  const fiscalYears = collectFiscalYearsInRange(start, end);
  let recordCount = 0;

  fiscalYears.forEach(function(fiscalYear) {
    ensureDefaultCyclePlanForFiscalYear(fiscalYear);
    const buildResult = buildDailyPlanFromCyclePlan(fiscalYear, end);
    recordCount += buildResult.dailyPlanCount;
  });

  upsertModuleSettingsValues(null, {
    PLAN_START_DATE: start,
    PLAN_END_DATE: end
  });

  syncModuleHoursWithCumulative(end, {
    preservePlanningRange: {
      startDate: start,
      endDate: end
    }
  });

  return {
    startDate: start,
    endDate: end,
    fiscalYears: fiscalYears,
    generatedAt: new Date(),
    recordCount: recordCount
  };
}

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
 * 日次計画と例外を合算して学年別合計を生成
 * @param {Object} dailyTotalsByGrade - 日次計画合計
 * @param {Object} exceptionTotals - 例外合計
 * @return {Object} 学年別合計
 */
function buildGradeTotalsFromDailyAndExceptions(dailyTotalsByGrade, exceptionTotals) {
  const result = createGradeTotalsTemplate();

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const daily = dailyTotalsByGrade[grade] || { plannedSessions: 0, elapsedSessions: 0, thisWeekSessions: 0 };
    const delta = toNumberOrZero(exceptionTotals.byGrade[grade]);
    const weeklyDelta = toNumberOrZero(exceptionTotals.thisWeekByGrade[grade]);
    const actual = Math.max(daily.elapsedSessions + delta, 0);

    result[grade].plannedSessions = daily.plannedSessions;
    result[grade].elapsedPlannedSessions = daily.elapsedSessions;
    result[grade].deltaSessions = delta;
    result[grade].actualSessions = actual;
    result[grade].diffSessions = actual - daily.elapsedSessions;
    result[grade].thisWeekSessions = daily.thisWeekSessions + weeklyDelta;
  }

  return result;
}

/**
 * 指定年度のデフォルトクール計画を必要時に作成
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @return {boolean} 作成した場合true
 */
function ensureDefaultCyclePlanForFiscalYear(fiscalYear, controlSheet) {
  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded().controlSheet;
  const existingRows = readCyclePlanRowsByFiscalYear(sheet, fiscalYear);

  if (existingRows.length > 0) {
    return false;
  }

  const rows = MODULE_DEFAULT_CYCLES.map(function(cycle) {
    return [
      Number(fiscalYear),
      cycle.order,
      cycle.startMonth,
      cycle.endMonth,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      MODULE_DEFAULT_KOMA_PER_CYCLE,
      cycle.label + ' default'
    ];
  });

  appendCyclePlanRows(sheet, rows);
  return true;
}

/**
 * 指定年度のクール計画を読み込み
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @return {Array<Object>} クール計画
 */
function loadCyclePlanForFiscalYear(fiscalYear, controlSheet) {
  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded().controlSheet;
  let rows = readCyclePlanRowsByFiscalYear(sheet, fiscalYear);

  if (rows.length === 0) {
    ensureDefaultCyclePlanForFiscalYear(fiscalYear, sheet);
    rows = readCyclePlanRowsByFiscalYear(sheet, fiscalYear);
  }

  const plans = rows.map(function(row) {
    return {
      fiscalYear: Number(fiscalYear),
      cycleOrder: Number(row[1]),
      startMonth: Number(row[2]),
      endMonth: Number(row[3]),
      gradeKoma: {
        1: Math.max(0, Math.round(toNumberOrZero(row[4]))),
        2: Math.max(0, Math.round(toNumberOrZero(row[5]))),
        3: Math.max(0, Math.round(toNumberOrZero(row[6]))),
        4: Math.max(0, Math.round(toNumberOrZero(row[7]))),
        5: Math.max(0, Math.round(toNumberOrZero(row[8]))),
        6: Math.max(0, Math.round(toNumberOrZero(row[9])))
      },
      note: row[10] || ''
    };
  }).filter(function(plan) {
    return Number.isInteger(plan.cycleOrder) &&
      Number.isInteger(plan.startMonth) &&
      Number.isInteger(plan.endMonth) &&
      plan.startMonth >= 1 && plan.startMonth <= 12 &&
      plan.endMonth >= 1 && plan.endMonth <= 12;
  });

  plans.sort(function(a, b) {
    return a.cycleOrder - b.cycleOrder;
  });

  if (plans.length === 0) {
    throw new Error('有効なクール計画がありません。モジュール学習管理画面で計画を確認してください。');
  }

  return plans;
}

/**
 * module_control から指定年度の計画行を抽出
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 対象年度
 * @return {Array<Array<*>>} 行データ
 */
function readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear) {
  return readAllCyclePlanRows(controlSheet).filter(function(row) {
    return Number(row[0]) === Number(fiscalYear);
  });
}

/**
 * 計画行を全件取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @return {Array<Array<*>>} 行データ
 */
function readAllCyclePlanRows(controlSheet) {
  const layout = getModuleControlLayout(controlSheet);
  const rowCount = layout.exceptionsMarkerRow - layout.planDataStartRow;
  if (rowCount <= 0) {
    return [];
  }

  const values = controlSheet.getRange(layout.planDataStartRow, 1, rowCount, MODULE_CONTROL_PLAN_HEADERS.length).getValues();
  return values.filter(function(row) {
    return row.some(function(value) {
      return isNonEmptyCell(value);
    });
  });
}

/**
 * クール計画から日次計画を構築（保存はしない）
 * @param {number} fiscalYear - 対象年度
 * @param {Date|string} baseDate - 集計基準日
 * @return {Object} 構築結果
 */
function buildDailyPlanFromCyclePlan(fiscalYear, baseDate) {
  return buildDailyPlanFromCyclePlanInternal(fiscalYear, baseDate, false);
}

/**
 * クール計画から日次計画を構築（内部実装）
 * @param {number} fiscalYear - 対象年度
 * @param {Date|string} baseDate - 集計基準日
 * @param {boolean} persistSheets - 互換引数（未使用）
 * @param {?Object} options - 実行オプション
 * @return {Object} 構築結果
 */
function buildDailyPlanFromCyclePlanInternal(fiscalYear, baseDate, persistSheets, options) {
  const normalizedFiscalYear = Number(fiscalYear);
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const generatedAt = new Date();
  const fiscalRange = getFiscalYearDateRange(normalizedFiscalYear);
  const weekStart = getWeekStartMonday(cutoffDate);
  const controlSheet = options && options.controlSheet ? options.controlSheet : null;
  const cyclePlans = loadCyclePlanForFiscalYear(normalizedFiscalYear, controlSheet);
  const schoolDayMap = buildSchoolDayMapByGradeForFiscalYear(normalizedFiscalYear);

  const dailyEntries = [];
  const planRows = [];
  const totalsByGrade = {};

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    totalsByGrade[grade] = {
      plannedSessions: 0,
      elapsedSessions: 0,
      thisWeekSessions: 0
    };
  }

  cyclePlans.forEach(function(plan) {
    const cycleLabel = plan.startMonth + '-' + plan.endMonth;
    const cycleMonthSet = buildCycleMonthKeySetForFiscalYear(normalizedFiscalYear, plan.startMonth, plan.endMonth);

    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const plannedKoma = toNumberOrZero(plan.gradeKoma[grade]);
      const plannedSessions = Math.max(0, Math.round(plannedKoma * 3));
      const gradeDates = schoolDayMap[grade].filter(function(date) {
        return !!cycleMonthSet[formatMonthKey(date)];
      });

      const weekMap = buildWeekMapFromDates(gradeDates);
      const allocations = allocateSessionsToDateKeys(plannedSessions, weekMap);
      const allocatedDateKeys = Object.keys(allocations).sort();

      allocatedDateKeys.forEach(function(dateKey) {
        const dateObj = normalizeToDate(dateKey);
        const sessions = allocations[dateKey];
        const elapsedFlag = dateObj <= cutoffDate ? 1 : 0;

        dailyEntries.push({
          date: dateObj,
          fiscalYear: normalizedFiscalYear,
          cycleOrder: plan.cycleOrder,
          cycleLabel: cycleLabel,
          weekKey: getWeekKey(dateObj),
          grade: grade,
          plannedSessions: sessions,
          elapsedFlag: elapsedFlag,
          generatedAt: generatedAt
        });

        totalsByGrade[grade].plannedSessions += sessions;
        if (elapsedFlag === 1) {
          totalsByGrade[grade].elapsedSessions += sessions;
          if (dateObj >= weekStart && dateObj <= cutoffDate) {
            totalsByGrade[grade].thisWeekSessions += sessions;
          }
        }
      });

      if (plannedSessions > 0 && allocatedDateKeys.length === 0) {
        Logger.log('[WARNING] 学校週が存在しないため割当をスキップしました: FY' + normalizedFiscalYear + ', cycle=' + cycleLabel + ', grade=' + grade);
      }

      planRows.push([
        normalizedFiscalYear,
        plan.cycleOrder,
        cycleLabel,
        grade,
        plannedKoma,
        plannedSessions,
        allocatedDateKeys.length,
        generatedAt
      ]);
    }
  });

  dailyEntries.sort(function(a, b) {
    if (a.date.getTime() !== b.date.getTime()) {
      return a.date.getTime() - b.date.getTime();
    }
    if (a.grade !== b.grade) {
      return a.grade - b.grade;
    }
    return a.cycleOrder - b.cycleOrder;
  });

  const dailyRows = dailyEntries.map(function(entry) {
    return [
      entry.date,
      entry.fiscalYear,
      entry.cycleOrder,
      entry.cycleLabel,
      entry.weekKey,
      entry.grade,
      entry.plannedSessions,
      entry.elapsedFlag,
      entry.generatedAt
    ];
  });

  if (persistSheets) {
    Logger.log('[INFO] 日次計画は module_control 運用のためシート保存をスキップしました。');
  }

  return {
    fiscalYear: normalizedFiscalYear,
    startDate: fiscalRange.startDate,
    endDate: fiscalRange.endDate,
    generatedAt: generatedAt,
    dailyPlanCount: dailyRows.length,
    dailyRows: dailyRows,
    planRows: planRows,
    totalsByGrade: totalsByGrade
  };
}

/**
 * 年度・学年別の学校日マップを構築
 * @param {number} fiscalYear - 対象年度
 * @return {Object} 学年別日付配列
 */
function buildSchoolDayMapByGradeForFiscalYear(fiscalYear) {
  const fiscalRange = getFiscalYearDateRange(fiscalYear);
  const result = {};

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    result[grade] = [];
  }

  const rows = extractSchoolDayRows(fiscalRange.startDate, fiscalRange.endDate);
  const unique = {};

  rows.forEach(function(row) {
    const date = row.date;
    const grade = row.grade;
    const dateKey = formatInputDate(date);
    const key = dateKey + '_' + grade;

    if (unique[key]) {
      return;
    }
    unique[key] = true;
    result[grade].push(date);
  });

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    result[grade].sort(function(a, b) {
      return a.getTime() - b.getTime();
    });
  }

  return result;
}

/**
 * 年間行事予定表から学校日候補を抽出
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array<Object>} 日付・学年配列
 */
function extractSchoolDayRows(startDate, endDate) {
  const sheet = getAnnualScheduleSheetOrThrow();
  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);

    if (!date || date < startDate || date > endDate) {
      continue;
    }

    const day = date.getDay();
    if (day === 0 || day === 6) {
      continue;
    }

    const grade = Number(row[SCHEDULE_COLUMNS.GRADE]);
    if (!Number.isInteger(grade) || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
      continue;
    }

    let hasSchoolData = false;
    for (let col = SCHEDULE_COLUMNS.DATA_START; col <= SCHEDULE_COLUMNS.DATA_END; col++) {
      if (isNonEmptyCell(row[col])) {
        hasSchoolData = true;
        break;
      }
    }

    if (!hasSchoolData) {
      continue;
    }

    rows.push({ date: date, grade: grade });
  }

  return rows;
}

/**
 * クール対象月の monthKey セットを生成
 * @param {number} fiscalYear - 年度
 * @param {number} startMonth - 開始月
 * @param {number} endMonth - 終了月
 * @return {Object} monthKeyセット
 */
function buildCycleMonthKeySetForFiscalYear(fiscalYear, startMonth, endMonth) {
  const months = listCycleMonthsInclusive(startMonth, endMonth);
  const set = {};

  months.forEach(function(month) {
    const year = month >= MODULE_FISCAL_YEAR_START_MONTH ? fiscalYear : fiscalYear + 1;
    set[year + '-' + String(month).padStart(2, '0')] = true;
  });

  return set;
}

/**
 * クール対象月（開始～終了、循環許容）
 * @param {number} startMonth - 開始月
 * @param {number} endMonth - 終了月
 * @return {Array<number>} 月配列
 */
function listCycleMonthsInclusive(startMonth, endMonth) {
  if (!Number.isInteger(startMonth) || !Number.isInteger(endMonth) ||
      startMonth < 1 || startMonth > 12 || endMonth < 1 || endMonth > 12) {
    return [];
  }

  const months = [];
  let cursor = startMonth;
  let guard = 0;

  while (guard < 13) {
    months.push(cursor);
    if (cursor === endMonth) {
      break;
    }
    cursor = cursor === 12 ? 1 : cursor + 1;
    guard++;
  }

  return months;
}

/**
 * セッションを学校週へ均等配分し、週内優先曜日で日付割当
 * @param {number} totalSessions - クール総セッション
 * @param {Object} weekMap - 週キーごとの学校日配列
 * @return {Object} dateKey別セッション数
 */
function allocateSessionsToDateKeys(totalSessions, weekMap) {
  const allocations = {};
  const weekKeys = Object.keys(weekMap).sort();

  if (totalSessions <= 0 || weekKeys.length === 0) {
    return allocations;
  }

  const basePerWeek = Math.floor(totalSessions / weekKeys.length);
  const remainder = totalSessions % weekKeys.length;

  weekKeys.forEach(function(weekKey, index) {
    const weekSessions = basePerWeek + (index < remainder ? 1 : 0);
    const orderedDates = sortWeekDatesByPriority(weekMap[weekKey]);

    if (weekSessions <= 0 || orderedDates.length === 0) {
      return;
    }

    for (let i = 0; i < weekSessions; i++) {
      const targetDate = orderedDates[i % orderedDates.length];
      const dateKey = formatInputDate(targetDate);
      allocations[dateKey] = toNumberOrZero(allocations[dateKey]) + 1;
    }
  });

  return allocations;
}

/**
 * 学校日を週単位にグルーピング
 * @param {Array<Date>} dates - 日付配列
 * @return {Object} 週キー => 日付配列
 */
function buildWeekMapFromDates(dates) {
  const map = {};

  dates.forEach(function(date) {
    const key = getWeekKey(date);
    if (!map[key]) {
      map[key] = [];
    }
    map[key].push(date);
  });

  return map;
}

/**
 * 週内日付を優先曜日順でソート
 * @param {Array<Date>} dates - 週内日付配列
 * @return {Array<Date>} ソート後
 */
function sortWeekDatesByPriority(dates) {
  return dates.slice().sort(function(a, b) {
    const priorityA = weekdayPriority(a.getDay());
    const priorityB = weekdayPriority(b.getDay());

    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }
    return a.getTime() - b.getTime();
  });
}

/**
 * 曜日優先度（月→水→金→火→木）
 * @param {number} dayOfWeek - Date#getDay()
 * @return {number} 優先度
 */
function weekdayPriority(dayOfWeek) {
  if (Object.prototype.hasOwnProperty.call(MODULE_WEEKDAY_PRIORITY, dayOfWeek)) {
    return MODULE_WEEKDAY_PRIORITY[dayOfWeek];
  }
  return 99;
}

/**
 * 週キー（週の月曜日）
 * @param {Date} date - 対象日
 * @return {string} yyyy-MM-dd
 */
function getWeekKey(date) {
  return formatInputDate(getWeekStartMonday(date));
}

/**
 * 対象日を含む週の月曜日を取得
 * @param {Date|string} value - 対象日
 * @return {Date} 月曜日
 */
function getWeekStartMonday(value) {
  const date = normalizeToDate(value);
  if (!date) {
    return null;
  }

  const copy = new Date(date.getTime());
  const day = copy.getDay();
  const shift = day === 0 ? -6 : (1 - day);
  copy.setDate(copy.getDate() + shift);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

/**
 * module_control の例外を集計（基準日まで）
 * @param {number} fiscalYear - 対象年度
 * @param {Date} baseDate - 基準日
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @return {Object} 学年別例外合計
 */
function loadExceptionTotals(fiscalYear, baseDate, controlSheet) {
  const totals = {
    byGrade: {},
    thisWeekByGrade: {}
  };
  const weekStart = getWeekStartMonday(baseDate);

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    totals.byGrade[grade] = 0;
    totals.thisWeekByGrade[grade] = 0;
  }

  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded().controlSheet;
  const rows = readExceptionRows(sheet);

  rows.forEach(function(item) {
    const exceptionDate = normalizeToDate(item.date);
    const grade = Number(item.grade);
    const delta = toNumberOrZero(item.deltaSessions);

    if (!exceptionDate || exceptionDate > baseDate) {
      return;
    }
    if (getFiscalYear(exceptionDate) !== fiscalYear) {
      return;
    }
    if (!Number.isInteger(grade) || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
      Logger.log('[WARNING] module_control の例外入力不正をスキップしました（行: ' + item.rowNumber + '）');
      return;
    }

    totals.byGrade[grade] += delta;
    if (exceptionDate >= weekStart && exceptionDate <= baseDate) {
      totals.thisWeekByGrade[grade] += delta;
    }
  });

  return totals;
}

/**
 * モジュール管理用シートを初期化
 * @return {Object} シート参照
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

  // 旧互換: 既存コードが参照するプロパティ名を残す
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
  let planMarkerRow = findMarkerRow(controlSheet, MODULE_CONTROL_MARKERS.PLAN, false);
  if (planMarkerRow < 1) {
    planMarkerRow = MODULE_CONTROL_DEFAULT_LAYOUT.PLAN_MARKER_ROW;
    controlSheet.getRange(planMarkerRow, 1).setValue(MODULE_CONTROL_MARKERS.PLAN);
  }

  let exceptionsMarkerRow = findMarkerRow(controlSheet, MODULE_CONTROL_MARKERS.EXCEPTIONS, true);
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
 * @return {Array<Object>} 例外行
 */
function readExceptionRows(controlSheet) {
  const layout = getModuleControlLayout(controlSheet);
  const lastRow = controlSheet.getLastRow();

  if (lastRow < layout.exceptionsDataStartRow) {
    return [];
  }

  const values = controlSheet
    .getRange(layout.exceptionsDataStartRow, 1, lastRow - layout.exceptionsDataStartRow + 1, MODULE_CONTROL_EXCEPTION_HEADERS.length)
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
      rowNumber: layout.exceptionsDataStartRow + index,
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
 * @return {number} 行数
 */
function countCyclePlanRowsForFiscalYear(controlSheet, fiscalYear) {
  return readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear).length;
}

/**
 * 年度別の例外行数をカウント
 * @param {GoogleAppsScript.Spreadsheet.Sheet} controlSheet - module_control
 * @param {number} fiscalYear - 年度
 * @return {number} 行数
 */
function countExceptionRowsForFiscalYear(controlSheet, fiscalYear) {
  return readExceptionRows(controlSheet).filter(function(item) {
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
    MODULE_SETTING_KEYS.DATA_VERSION,
    MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN
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
 * 旧互換: データ移行関数（初期化時に実行済み）
 */
function migrateModuleDataIfNeeded() {
  return;
}

/**
 * 旧互換: 例外シート移行関数（初期化時に実行済み）
 */
function migrateModuleExceptionsSheetIfNeeded() {
  return;
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
  const cumulativeSheet = getSheetByNameOrThrow('累計時数');

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
 * @return {number} 列番号（1-based）
 */
function resolveCumulativeDisplayColumn(cumulativeSheet) {
  const settingsMap = readModuleSettingsMap();
  const displayRowCount = MODULE_GRADE_MAX - MODULE_GRADE_MIN + 1;
  const preferredColumn = MODULE_CUMULATIVE_COLUMNS.DISPLAY_FALLBACK;

  const preferredHeader = String(cumulativeSheet.getRange(2, preferredColumn).getValue() || '').trim();
  if (preferredHeader === MODULE_DISPLAY_HEADER ||
      (preferredHeader === '' &&
        isReusableCumulativeDisplayColumn(cumulativeSheet, preferredColumn, displayRowCount))) {
    cumulativeSheet.getRange(2, preferredColumn).setValue(MODULE_DISPLAY_HEADER);
    return preferredColumn;
  }

  const configuredColumn = Number(settingsMap[MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN]);
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
 * @return {Object} 期間
 */
function getModulePlanningRangeFromSettings(settingsSheet, fallbackDate) {
  const map = readModuleSettingsMap();
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
 * 旧API互換: 学校日ベース月次マップを返却
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
  const fiscalYears = collectFiscalYearsInRange(start, end);
  const countedDays = {};

  fiscalYears.forEach(function(fiscalYear) {
    ensureDefaultCyclePlanForFiscalYear(fiscalYear);
    const buildResult = buildDailyPlanFromCyclePlanInternal(fiscalYear, end, false);

    buildResult.dailyRows.forEach(function(row) {
      const date = normalizeToDate(row[0]);
      const grade = Number(row[5]);
      const sessions = toNumberOrZero(row[6]);

      if (!date || date < start || date > end) {
        return;
      }
      if (!Number.isInteger(grade) || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
        return;
      }

      const monthKey = formatMonthKey(date);
      const entry = planMap.byMonth[monthKey] && planMap.byMonth[monthKey][grade];
      if (!entry) {
        return;
      }

      entry.planned_sessions += sessions;

      const dayKey = monthKey + '_' + grade + '_' + formatInputDate(date);
      if (!countedDays[dayKey]) {
        countedDays[dayKey] = true;
        entry.school_days_count += 1;
      }
    });
  });

  Object.keys(planMap.byMonth).forEach(function(monthKey) {
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      entry.planned_units = sessionsToUnits(entry.planned_sessions);
      entry.actual_units = entry.planned_units;
      entry.diff_units = 0;
    }
  });

  return planMap;
}

/**
 * 旧API互換: 例外差分を月次マップへ反映
 * @param {Object} planMap - 計画マップ
 * @param {Date|string} baseDate - 基準日
 * @return {Object} 差分反映後
 */
function applyModuleExceptions(planMap, baseDate) {
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const rows = readExceptionRows(sheets.controlSheet);

  Object.keys(planMap.byMonth).forEach(function(monthKey) {
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      entry.delta_units = 0;
      entry.actual_units = entry.planned_units;
      entry.diff_units = 0;
    }
  });

  rows.forEach(function(item) {
    const exceptionDate = normalizeToDate(item.date);
    const grade = Number(item.grade);
    const deltaSessions = toNumberOrZero(item.deltaSessions);

    if (!exceptionDate || exceptionDate > cutoffDate) {
      return;
    }
    if (!Number.isInteger(grade) || grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
      Logger.log('[WARNING] module_control の例外入力不正をスキップしました（行: ' + item.rowNumber + '）');
      return;
    }

    const monthKey = formatMonthKey(exceptionDate);
    if (!planMap.byMonth[monthKey] || !planMap.byMonth[monthKey][grade]) {
      return;
    }

    planMap.byMonth[monthKey][grade].delta_units += sessionsToUnits(deltaSessions);
  });

  Object.keys(planMap.byMonth).forEach(function(monthKey) {
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const entry = planMap.byMonth[monthKey][grade];
      entry.actual_units = Math.max(entry.planned_units + entry.delta_units, 0);
      entry.diff_units = entry.actual_units - entry.planned_units;
    }
  });

  return planMap;
}

/**
 * 空の計画マップを作成（旧互換）
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 計画マップ
 */
function createEmptyPlanMap(startDate, endDate) {
  const map = { byMonth: {} };
  const monthKeys = listMonthKeysInRange(startDate, endDate);

  monthKeys.forEach(function(monthKey) {
    map.byMonth[monthKey] = {};
    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      map.byMonth[monthKey][grade] = {
        planned_sessions: 0,
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
 * 月キー一覧を生成
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array<string>} 月キー
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
