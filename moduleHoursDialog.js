/**
 * @fileoverview モジュール学習管理 - ダイアログ/ユーザー操作
 * @description モジュール学習管理のダイアログ表示、ユーザー入力処理を担当します。
 */

/**
 * モジュール学習管理ダイアログを表示
 */
function showModulePlanningDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('modulePlanningDialog').evaluate()
      .setWidth(980)
      .setHeight(800);
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
  const enabledWeekdays = getEnabledWeekdays(settingsMap);
  const savedRange = getModulePlanningRangeFromSettings(baseDate, settingsMap);
  const annualTarget = buildDialogAnnualTargetForFiscalYear(fiscalYear, controlSheet, annualTargetRows);
  const recentExceptions = listRecentExceptionsForFiscalYear(controlSheet, fiscalYear, 10, exceptionRows);
  const annualTargetRecordCount = countAnnualTargetRowsForFiscalYear(controlSheet, fiscalYear, annualTargetRows);
  const exceptionRecordCount = countExceptionRowsForFiscalYear(controlSheet, fiscalYear, exceptionRows);
  const cumulativeDisplayColumn = String(MODULE_CUMULATIVE_COLUMNS.DISPLAY);

  // 予備セッション数・日次件数をリアルタイム算出（実施期間を反映）
  const buildResult = buildDailyPlanFromAnnualTarget(fiscalYear, baseDate, {
    controlSheet: controlSheet,
    enabledWeekdays: enabledWeekdays,
    startDate: savedRange.startDate,
    endDate: savedRange.endDate
  });
  const reserveByGrade = buildResult.reserveByGrade;
  const dailyPlanCount = buildResult.dailyPlanCount;

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
    defaultExceptionDate: formatInputDate(getDefaultExceptionDate(enabledWeekdays)),
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
    enabledWeekdays: enabledWeekdays,
    weekdayLabels: MODULE_WEEKDAY_LABELS,
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
 * 差分入力のデフォルト日付を取得（直近の実施曜日）
 * @param {Array<number>} enabledWeekdays - 有効曜日配列（getDay()値）
 * @return {Date} デフォルト日付
 */
function getDefaultExceptionDate(enabledWeekdays) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = today.getDay();

  if (Array.isArray(enabledWeekdays) && enabledWeekdays.indexOf(day) !== -1) {
    return today;
  }

  for (let offset = 1; offset <= 7; offset++) {
    const candidate = new Date(today.getTime());
    candidate.setDate(candidate.getDate() - offset);
    if (Array.isArray(enabledWeekdays) && enabledWeekdays.indexOf(candidate.getDay()) !== -1) {
      return candidate;
    }
  }

  return today;
}

/**
 * ダイアログ表示用の年間目標を取得（V4形式）
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @param {Array<Array<*>>=} annualTargetRows - 事前取得済み年間目標行（対象年度）
 * @return {Object} 年間目標（grades, note）
 */
function buildDialogAnnualTargetForFiscalYear(fiscalYear, controlSheet, annualTargetRows) {
  const target = Array.isArray(annualTargetRows) && annualTargetRows.length > 0
    ? buildAnnualTargetFromRows(fiscalYear, annualTargetRows)
    : loadAnnualTargetForFiscalYear(fiscalYear, controlSheet);

  const note = Array.isArray(annualTargetRows) && annualTargetRows.length > 0
    ? String(annualTargetRows[0][16] || '').trim()
    : '';

  return {
    grades: target.grades,
    note: note
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
  const rows = normalizeAnnualTargetRowsFromDialog(fiscalYear, target);

  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  replaceAnnualTargetRowsForFiscalYearInControl(controlSheet, fiscalYear, rows);

  const baseDate = normalizeToDate(payload && payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const result = syncModuleHoursWithCumulative(baseDate);

  const lines = [
    '年間目標を保存して再集計しました。',
    '対象年度: ' + fiscalYear + '年度',
    '基準日: ' + formatInputDate(result.baseDate)
  ];

  const deficitWarning = buildDeficitWarningMessage(result.reserveByGrade);
  if (deficitWarning) {
    lines.push('');
    lines.push(deficitWarning);
  }

  return lines.join('\n');
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

  const dayOfWeek = exceptionDate.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    throw new Error(formatInputDate(exceptionDate) + ' は土日です。実施日（平日）を指定してください。');
  }

  const enabledWeekdays = getEnabledWeekdays();
  if (enabledWeekdays.indexOf(dayOfWeek) === -1) {
    const dayLabel = MODULE_WEEKDAY_LABELS[dayOfWeek] || '';
    const enabledLabels = enabledWeekdays
      .slice().sort(function(a, b) { return a - b; })
      .map(function(d) { return MODULE_WEEKDAY_LABELS[d] || String(d); })
      .join('・');
    throw new Error(formatInputDate(exceptionDate) + '（' + dayLabel + '）は実施曜日ではありません。実施曜日: ' + enabledLabels);
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

  const minuteSign = deltaSessions > 0 ? '+' : '';
  return [
    '実施差分を保存して再集計しました。',
    '入力: ' + formatInputDate(exceptionDate) + ' / ' + grade + '年 / ' +
      formatSignedSessionsAsMixedFraction(deltaSessions) + 'コマ（' + minuteSign + (deltaSessions * 15) + '分）',
    '基準日: ' + formatInputDate(result.baseDate)
  ].join('\n');
}

/**
 * ダイアログ入力値をV4形式の年間目標行群へ正規化
 * @param {number} fiscalYear - 対象年度
 * @param {Object} target - 入力目標 { grades: {grade: {mode, annualKoma, monthlyKoma}}, note }
 * @return {Array<Array<*>>} シート行群（MODULE_CONTROL_PLAN_HEADERS形式 × 学年数）
 */
function normalizeAnnualTargetRowsFromDialog(fiscalYear, target) {
  if (!target || !target.grades) {
    throw new Error('年間目標のデータがありません。');
  }

  const note = String(target.note || '').trim();
  const rows = [];

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const gradeData = target.grades[grade] || {};
    const mode = gradeData.mode === MODULE_PLAN_MODE_MONTHLY
      ? MODULE_PLAN_MODE_MONTHLY
      : MODULE_PLAN_MODE_ANNUAL;

    let annualKoma;
    let monthlyKoma = null;

    if (mode === MODULE_PLAN_MODE_MONTHLY) {
      monthlyKoma = {};
      let monthlyTotal = 0;
      [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].forEach(function(m) {
        const val = Math.max(0, Math.round(toNumberOrZero(gradeData.monthlyKoma && gradeData.monthlyKoma[m])));
        monthlyKoma[m] = val;
        monthlyTotal += val;
      });
      if (monthlyTotal <= 0) {
        throw new Error(grade + '年: 月別モードの合計コマ数が0です。');
      }
      annualKoma = monthlyTotal;
    } else {
      annualKoma = Math.round(toNumberOrZero(gradeData.annualKoma));
      if (annualKoma < 0) {
        throw new Error(grade + '年のコマ数は0以上で入力してください。');
      }
    }

    rows.push(buildV4PlanRow(fiscalYear, grade, mode, annualKoma, monthlyKoma, note));
  }

  return rows;
}

/**
 * ダイアログから実施設定（実施曜日・実施期間）を保存して再集計
 * @param {Object} payload - { weekdays: [1,3,5], startDate: 'yyyy-MM-dd', endDate: 'yyyy-MM-dd', baseDate: ... }
 * @return {string} 完了メッセージ
 */
function saveModuleSettingsFromDialog(payload) {
  if (!payload) {
    throw new Error('設定データがありません。');
  }

  const weekdays = Array.isArray(payload.weekdays) ? payload.weekdays : [];
  const validWeekdays = weekdays
    .map(function(d) { return parseInt(d, 10); })
    .filter(function(n) { return Number.isInteger(n) && n >= 1 && n <= 5; });

  if (validWeekdays.length === 0) {
    throw new Error('実施曜日を1つ以上選択してください。');
  }

  const startDate = normalizeToDate(payload.startDate);
  const endDate = normalizeToDate(payload.endDate);
  if (!startDate || !endDate) {
    throw new Error('実施期間の開始日・終了日を正しく入力してください。');
  }
  if (startDate > endDate) {
    throw new Error('開始日は終了日以前の日付を指定してください。');
  }

  const baseDate = normalizeToDate(payload.baseDate) || normalizeToDate(getCurrentOrNextSaturday());
  const fiscalYear = getFiscalYear(baseDate);
  const fiscalRange = getFiscalYearDateRange(fiscalYear);
  if (startDate < fiscalRange.startDate || endDate > fiscalRange.endDate) {
    throw new Error('実施期間は年度範囲内（' + formatInputDate(fiscalRange.startDate) + ' ～ ' + formatInputDate(fiscalRange.endDate) + '）で指定してください。');
  }

  upsertModuleSettingsValues({
    WEEKDAYS_ENABLED: validWeekdays,
    PLAN_START_DATE: startDate,
    PLAN_END_DATE: endDate
  });

  const result = syncModuleHoursWithCumulative(baseDate, {
    preservePlanningRange: {
      startDate: startDate,
      endDate: endDate
    }
  });

  const dayNames = validWeekdays
    .sort(function(a, b) { return a - b; })
    .map(function(d) { return MODULE_WEEKDAY_LABELS[d] || String(d); })
    .join('・');

  const lines = [
    '実施設定を保存して再集計しました。',
    '実施曜日: ' + dayNames,
    '実施期間: ' + formatInputDate(startDate) + ' ～ ' + formatInputDate(endDate),
    '基準日: ' + formatInputDate(result.baseDate)
  ];

  const deficitWarning = buildDeficitWarningMessage(result.reserveByGrade);
  if (deficitWarning) {
    lines.push('');
    lines.push(deficitWarning);
  }

  return lines.join('\n');
}

/**
 * 不足学年の警告メッセージを生成
 * @param {Object} reserveByGrade - 学年別予備セッション数
 * @return {string} 不足がなければ空文字列
 */
function buildDeficitWarningMessage(reserveByGrade) {
  if (!reserveByGrade) {
    return '';
  }

  const deficits = [];
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const reserve = toNumberOrZero(reserveByGrade[grade]);
    if (reserve < 0) {
      deficits.push(grade + '年: ' + MODULE_DEFICIT_LABEL + ' ' + formatSessionsAsMixedFraction(Math.abs(reserve)) + 'コマ');
    }
  }

  if (deficits.length === 0) {
    return '';
  }

  return '【注意】実施可能日数に対して目標コマ数が不足しています。\n' + deficits.join('、');
}

