/**
 * @fileoverview モジュール学習管理 - ダイアログ/ユーザー操作
 * @description モジュール学習管理のダイアログ表示、ユーザー入力処理を担当します。
 */

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
  let initElapsedMs = 0;
  let dataElapsedMs = 0;
  let cumulativeElapsedMs = 0;

  const initStartedAt = new Date().getTime();
  const sheets = initializeModuleHoursSheetsIfNeeded();
  initElapsedMs = new Date().getTime() - initStartedAt;

  const controlSheet = sheets.controlSheet;
  const baseDate = normalizeToDate(getCurrentOrNextSaturday());
  const fiscalYear = getFiscalYear(baseDate);
  const fiscalRange = getFiscalYearDateRange(fiscalYear);

  const dataStartedAt = new Date().getTime();
  let layout = getModuleControlLayout(controlSheet);
  let cyclePlanRows = readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear, null, layout);
  const createdDefaults = ensureDefaultCyclePlanForFiscalYear(fiscalYear, controlSheet, cyclePlanRows);
  if (createdDefaults) {
    layout = getModuleControlLayout(controlSheet);
    cyclePlanRows = readCyclePlanRowsByFiscalYear(controlSheet, fiscalYear, null, layout);
  }
  const exceptionRows = readExceptionRows(controlSheet, layout);
  dataElapsedMs = new Date().getTime() - dataStartedAt;

  const settingsMap = readModuleSettingsMap();
  const savedRange = getModulePlanningRangeFromSettings(null, baseDate, settingsMap);
  const dailyPlanCount = getCachedDailyPlanCountForDialog(settingsMap);
  const cyclePlans = buildDialogCyclePlansForFiscalYear(fiscalYear, controlSheet, cyclePlanRows);
  const recentExceptions = listRecentExceptionsForFiscalYear(controlSheet, fiscalYear, 10, exceptionRows);
  const cyclePlanRecordCount = countCyclePlanRowsForFiscalYear(controlSheet, fiscalYear, cyclePlanRows);
  const exceptionRecordCount = countExceptionRowsForFiscalYear(controlSheet, fiscalYear, exceptionRows);
  let cumulativeDisplayColumn = settingsMap[MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN] || '';

  const cumulativeStartedAt = new Date().getTime();
  try {
    const cumulativeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUMULATIVE_SHEET.NAME);
    if (cumulativeSheet) {
      const displayColumn = resolveCumulativeDisplayColumn(cumulativeSheet, settingsMap);
      enforceModuleCumulativeColumnVisibility(cumulativeSheet, displayColumn);
      cumulativeDisplayColumn = String(displayColumn);
    }
  } catch (error) {
    Logger.log('[WARNING] 累計時数の列表示制御に失敗: ' + error.toString());
  }
  cumulativeElapsedMs = new Date().getTime() - cumulativeStartedAt;

  const state = {
    baseDate: formatInputDate(baseDate),
    fiscalYear: fiscalYear,
    fiscalYearStartDate: formatInputDate(fiscalRange.startDate),
    fiscalYearEndDate: formatInputDate(fiscalRange.endDate),
    startDate: formatInputDate(savedRange.startDate),
    endDate: formatInputDate(savedRange.endDate),
    lastGeneratedAt: formatDateTimeForDisplay(settingsMap[MODULE_SETTING_KEYS.LAST_GENERATED_AT]),
    cyclePlanRecordCount: cyclePlanRecordCount,
    dailyPlanRecordCount: dailyPlanCount,
    exceptionRecordCount: exceptionRecordCount,
    cumulativeDisplayColumn: cumulativeDisplayColumn,
    cyclePlans: cyclePlans,
    recentExceptions: recentExceptions
  };

  const elapsedMs = new Date().getTime() - startedAt;
  if (elapsedMs >= 2000) {
    Logger.log('[PERF] getModulePlanningDialogState: ' + elapsedMs +
      'ms (init=' + initElapsedMs +
      'ms, data=' + dataElapsedMs +
      'ms, cumulative=' + cumulativeElapsedMs + 'ms)');
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
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @param {Array<Array<*>>=} cyclePlanRows - 事前取得済み計画行（対象年度）
 * @return {Array<Object>} 計画配列
 */
function buildDialogCyclePlansForFiscalYear(fiscalYear, controlSheet, cyclePlanRows) {
  const plans = Array.isArray(cyclePlanRows)
    ? toCyclePlansFromRows(fiscalYear, cyclePlanRows)
    : loadCyclePlanForFiscalYear(fiscalYear, controlSheet);
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
 * @param {Array<Object>=} exceptionRows - 事前取得済み例外行
 * @return {Array<Object>} 実施差分配列
 */
function listRecentExceptionsForFiscalYear(controlSheet, fiscalYear, limitCount, exceptionRows) {
  const limit = Math.max(1, Number(limitCount) || 10);
  const targetFiscalYear = Number(fiscalYear);
  const rows = Array.isArray(exceptionRows) ? exceptionRows : readExceptionRows(controlSheet);

  return rows
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
    'モジュール学習の再集計が完了しました。',
    '基準日: ' + formatInputDate(result.baseDate),
    '対象年度: ' + result.fiscalYear + '年度',
    '日次計画件数（再集計結果）: ' + result.dailyPlanCount + '件'
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
    'クール計画件数: ' + rows.length + '件',
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
    throw new Error('差分値は0以外の数値を入力してください。');
  }

  const reason = String(payload && payload.reason ? payload.reason : '').trim();
  const note = String(payload && payload.note ? payload.note : '').trim();

  const sheets = initializeModuleHoursSheetsIfNeeded();
  appendExceptionRows(sheets.controlSheet, [[exceptionDate, grade, deltaSessions, reason, note]]);

  const baseDate = normalizeToDate(payload && payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const result = syncModuleHoursWithCumulative(baseDate);

  return [
    '実施差分を保存して再集計しました。',
    '入力: ' + formatInputDate(exceptionDate) + ' / ' + grade + '年 / ' +
      formatSignedSessionsAsMixedFraction(deltaSessions) + 'コマ（' + (deltaSessions * 15) + '分）',
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
