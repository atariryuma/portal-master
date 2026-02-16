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
 * ダイアログ表示用の状態を返却
 * @return {Object} ダイアログ状態
 */
function getModulePlanningDialogState() {
  const startedAt = new Date().getTime();
  let initElapsedMs = 0;
  let dataElapsedMs = 0;
  let cumulativeElapsedMs = 0;

  const initStartedAt = new Date().getTime();
  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  initElapsedMs = new Date().getTime() - initStartedAt;
  const baseDate = normalizeToDate(getCurrentOrNextSaturday());
  const fiscalYear = getFiscalYear(baseDate);
  const fiscalRange = getFiscalYearDateRange(fiscalYear);

  const dataStartedAt = new Date().getTime();
  let layout = getModuleControlLayout(controlSheet);
  let annualTargetRows = readAnnualTargetRowsByFiscalYear(controlSheet, fiscalYear, null, layout);
  const createdDefaults = ensureDefaultAnnualTargetForFiscalYear(fiscalYear, controlSheet, annualTargetRows);
  if (createdDefaults) {
    layout = getModuleControlLayout(controlSheet);
    annualTargetRows = readAnnualTargetRowsByFiscalYear(controlSheet, fiscalYear, null, layout);
  }
  const exceptionRows = readExceptionRows(controlSheet, layout);
  dataElapsedMs = new Date().getTime() - dataStartedAt;

  const settingsMap = readModuleSettingsMap();
  const savedRange = getModulePlanningRangeFromSettings(baseDate, settingsMap);
  const dailyPlanCount = getCachedDailyPlanCountForDialog(settingsMap);
  const annualTarget = buildDialogAnnualTargetForFiscalYear(fiscalYear, controlSheet, annualTargetRows);
  const recentExceptions = listRecentExceptionsForFiscalYear(controlSheet, fiscalYear, 10, exceptionRows);
  const annualTargetRecordCount = countAnnualTargetRowsForFiscalYear(controlSheet, fiscalYear, annualTargetRows);
  const exceptionRecordCount = countExceptionRowsForFiscalYear(controlSheet, fiscalYear, exceptionRows);
  const cumulativeDisplayColumn = String(MODULE_CUMULATIVE_COLUMNS.DISPLAY);

  // 予備セッション数を算出
  const buildResult = buildDailyPlanFromAnnualTarget(fiscalYear, baseDate, {
    controlSheet: controlSheet
  });
  const reserveByGrade = buildResult.reserveByGrade;

  const cumulativeStartedAt = new Date().getTime();
  try {
    const cumulativeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUMULATIVE_SHEET.NAME);
    if (cumulativeSheet) {
      cumulativeSheet.hideColumns(MODULE_CUMULATIVE_COLUMNS.PLAN, 3);
      cumulativeSheet.showColumns(MODULE_CUMULATIVE_COLUMNS.DISPLAY, 1);
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
    annualTargetRecordCount: annualTargetRecordCount,
    dailyPlanRecordCount: dailyPlanCount,
    exceptionRecordCount: exceptionRecordCount,
    cumulativeDisplayColumn: cumulativeDisplayColumn,
    annualTarget: annualTarget,
    reserveByGrade: reserveByGrade,
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
 * ダイアログ表示用の年間目標を取得
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @param {Array<Array<*>>=} annualTargetRows - 事前取得済み年間目標行（対象年度）
 * @return {Object} 年間目標（g1Koma〜g6Koma, note）
 */
function buildDialogAnnualTargetForFiscalYear(fiscalYear, controlSheet, annualTargetRows) {
  const target = Array.isArray(annualTargetRows) && annualTargetRows.length > 0
    ? toAnnualTargetFromRow(fiscalYear, annualTargetRows[0])
    : loadAnnualTargetForFiscalYear(fiscalYear, controlSheet);
  return {
    g1Koma: target.gradeKoma[1],
    g2Koma: target.gradeKoma[2],
    g3Koma: target.gradeKoma[3],
    g4Koma: target.gradeKoma[4],
    g5Koma: target.gradeKoma[5],
    g6Koma: target.gradeKoma[6],
    note: target.note || ''
  };
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
 * ダイアログから受け取った年間目標を保存して再集計
 * @param {Object} payload - 入力データ
 * @return {string} 完了メッセージ
 */
function saveModuleAnnualTargetFromDialog(payload) {
  const fiscalYear = Number(payload && payload.fiscalYear);
  if (!Number.isInteger(fiscalYear) || fiscalYear < 2000 || fiscalYear > 2100) {
    throw new Error('対象年度が不正です。');
  }

  const target = payload && payload.target ? payload.target : null;
  const row = normalizeAnnualTargetRowFromDialog(fiscalYear, target);

  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  replaceAnnualTargetRowsForFiscalYearInControl(controlSheet, fiscalYear, [row]);

  const baseDate = normalizeToDate(payload && payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const result = syncModuleHoursWithCumulative(baseDate);

  return [
    '年間目標を保存して再集計しました。',
    '対象年度: ' + fiscalYear + '年度',
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

  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  appendExceptionRows(controlSheet, [[exceptionDate, grade, deltaSessions, reason, note]]);

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
 * ダイアログ入力値を年間目標行へ正規化
 * @param {number} fiscalYear - 対象年度
 * @param {Object} target - 入力目標 { g1Koma, ..., g6Koma, note }
 * @return {Array<*>} シート行（MODULE_CONTROL_PLAN_HEADERS形式）
 */
function normalizeAnnualTargetRowFromDialog(fiscalYear, target) {
  if (!target) {
    throw new Error('年間目標のデータがありません。');
  }

  const gradeValues = [];
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const key = 'g' + grade + 'Koma';
    const rawValue = toNumberOrZero(target[key]);
    if (rawValue < 0) {
      throw new Error(grade + '年のコマ数は0以上で入力してください。');
    }
    gradeValues.push(Math.round(rawValue));
  }

  const note = String(target.note ? target.note : '').trim();

  return [
    Number(fiscalYear),
    gradeValues[0],
    gradeValues[1],
    gradeValues[2],
    gradeValues[3],
    gradeValues[4],
    gradeValues[5],
    note
  ];
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
