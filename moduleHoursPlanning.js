/**
 * @fileoverview モジュール学習管理 - 配分アルゴリズム
 * @description 年間目標から日次配分の構築、学校日マップ、例外集計を担当します。
 */

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
 * 指定年度のデフォルト年間目標を必要時に作成
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @param {Array<Array<*>>=} existingRowsForFiscalYear - 事前取得済みの対象年度行
 * @return {boolean} 作成した場合true
 */
function ensureDefaultAnnualTargetForFiscalYear(fiscalYear, controlSheet, existingRowsForFiscalYear) {
  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded();
  const existingRows = Array.isArray(existingRowsForFiscalYear)
    ? existingRowsForFiscalYear
    : readAnnualTargetRowsByFiscalYear(sheet, fiscalYear);

  if (existingRows.length > 0) {
    return false;
  }

  const rows = [];
  for (let g = MODULE_GRADE_MIN; g <= MODULE_GRADE_MAX; g++) {
    rows.push(buildV4PlanRow(Number(fiscalYear), g, MODULE_PLAN_MODE_ANNUAL, MODULE_DEFAULT_ANNUAL_KOMA, null, 'default'));
  }

  appendAnnualTargetRows(sheet, rows);
  return true;
}

/**
 * 指定年度の年間目標を読み込み
 * @param {number} fiscalYear - 対象年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} controlSheet - module_control
 * @return {Object} 年間目標 { fiscalYear, grades: {grade: {mode, annualKoma, monthlyKoma}}, gradeKoma: {grade: N} }
 */
function loadAnnualTargetForFiscalYear(fiscalYear, controlSheet) {
  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded();
  let rows = readAnnualTargetRowsByFiscalYear(sheet, fiscalYear);

  if (rows.length === 0) {
    ensureDefaultAnnualTargetForFiscalYear(fiscalYear, sheet);
    rows = readAnnualTargetRowsByFiscalYear(sheet, fiscalYear);
  }

  if (rows.length === 0) {
    throw new Error('年間目標が取得できません（年度: ' + fiscalYear + '）');
  }

  return buildAnnualTargetFromRows(fiscalYear, rows);
}

/**
 * V4形式の複数行から年間目標オブジェクトを構築
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>} rows - 行データ（V4形式: 学年別行）
 * @return {Object} 年間目標
 */
function buildAnnualTargetFromRows(fiscalYear, rows) {
  const grades = {};
  const gradeKoma = {};

  for (let g = MODULE_GRADE_MIN; g <= MODULE_GRADE_MAX; g++) {
    grades[g] = { mode: MODULE_PLAN_MODE_ANNUAL, annualKoma: MODULE_DEFAULT_ANNUAL_KOMA, monthlyKoma: null };
    gradeKoma[g] = MODULE_DEFAULT_ANNUAL_KOMA;
  }

  rows.forEach(function(row) {
    const grade = Number(row[1]);
    if (grade < MODULE_GRADE_MIN || grade > MODULE_GRADE_MAX) {
      return;
    }
    const mode = String(row[2] || '').trim() === MODULE_PLAN_MODE_MONTHLY
      ? MODULE_PLAN_MODE_MONTHLY
      : MODULE_PLAN_MODE_ANNUAL;
    const annualKoma = Math.max(0, Math.round(toNumberOrZero(row[15])));

    if (mode === MODULE_PLAN_MODE_MONTHLY) {
      const monthlyKoma = {};
      [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].forEach(function(m, i) {
        monthlyKoma[m] = Math.max(0, Math.round(toNumberOrZero(row[3 + i])));
      });
      grades[grade] = { mode: MODULE_PLAN_MODE_MONTHLY, annualKoma: annualKoma, monthlyKoma: monthlyKoma };
    } else {
      grades[grade] = { mode: MODULE_PLAN_MODE_ANNUAL, annualKoma: annualKoma, monthlyKoma: null };
    }
    gradeKoma[grade] = annualKoma;
  });

  return {
    fiscalYear: Number(fiscalYear),
    grades: grades,
    gradeKoma: gradeKoma
  };
}

/**
 * 年間目標から日次計画を構築（保存はしない）
 * 実施期間内の実施可能日に対してセッションを均等配分し、予備セッション数も算出する。
 * options.startDate/endDate で実施期間を指定可能（省略時は年度全体）。
 * @param {number} fiscalYear - 対象年度
 * @param {Date|string} baseDate - 集計基準日
 * @param {?Object} options - 実行オプション
 * @param {Sheet} [options.controlSheet] - module_control シート（省略時は自動取得）
 * @param {number[]} [options.enabledWeekdays] - 有効曜日配列（省略時は設定値）
 * @param {Date|string} [options.startDate] - 実施開始日（省略時は年度開始日）
 * @param {Date|string} [options.endDate] - 実施終了日（省略時は年度終了日）
 * @return {Object} 構築結果（totalsByGrade, reserveByGrade, dailyPlanCount 含む）
 */
function buildDailyPlanFromAnnualTarget(fiscalYear, baseDate, options) {
  const normalizedFiscalYear = Number(fiscalYear);
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const generatedAt = new Date();
  const fiscalRange = getFiscalYearDateRange(normalizedFiscalYear);
  const weekStart = getWeekStartMonday(cutoffDate);
  const controlSheet = options && options.controlSheet ? options.controlSheet : null;
  const annualTarget = loadAnnualTargetForFiscalYear(normalizedFiscalYear, controlSheet);
  const enabledWeekdays = options && Array.isArray(options.enabledWeekdays)
    ? options.enabledWeekdays
    : getEnabledWeekdays();
  const planStartDate = options && options.startDate ? normalizeToDate(options.startDate) : null;
  const planEndDate = options && options.endDate ? normalizeToDate(options.endDate) : null;
  const schoolDayMap = buildSchoolDayMapByGradeForFiscalYear(normalizedFiscalYear, enabledWeekdays, planStartDate, planEndDate);

  const dailyEntries = [];
  const planRows = [];
  const totalsByGrade = {};
  const reserveByGrade = {};

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    totalsByGrade[grade] = {
      plannedSessions: 0,
      elapsedSessions: 0,
      thisWeekSessions: 0
    };
  }

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    const gradeTarget = annualTarget.grades[grade];
    const plannedKoma = toNumberOrZero(annualTarget.gradeKoma[grade]);
    const plannedSessions = Math.max(0, Math.round(plannedKoma * 3));
    const gradeDates = schoolDayMap[grade];

    // モード別にセッションを配分
    let allocations;
    if (gradeTarget.mode === MODULE_PLAN_MODE_MONTHLY && gradeTarget.monthlyKoma) {
      allocations = allocateSessionsByMonth(gradeTarget.monthlyKoma, gradeDates);
    } else {
      const weekMap = buildWeekMapFromDates(gradeDates);
      allocations = allocateSessionsToDateKeys(plannedSessions, weekMap);
    }
    const allocatedDateKeys = Object.keys(allocations).sort();

    // 予備/不足 = 実施可能日数 - 目標セッション数（正=予備、負=不足）
    // 1日1回上限により配分できない分も不足として扱う。
    reserveByGrade[grade] = gradeDates.length - plannedSessions;

    allocatedDateKeys.forEach(function(dateKey) {
      const dateObj = normalizeToDate(dateKey);
      const sessions = allocations[dateKey];
      const elapsedFlag = dateObj <= cutoffDate ? 1 : 0;

      dailyEntries.push({
        date: dateObj,
        fiscalYear: normalizedFiscalYear,
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
      Logger.log('[WARNING] 学校週が存在しないため割当をスキップしました: FY' + normalizedFiscalYear + ', grade=' + grade);
    }

    planRows.push([
      normalizedFiscalYear,
      0,
      '',
      grade,
      plannedKoma,
      plannedSessions,
      allocatedDateKeys.length,
      generatedAt
    ]);
  }

  dailyEntries.sort(function(a, b) {
    if (a.date.getTime() !== b.date.getTime()) {
      return a.date.getTime() - b.date.getTime();
    }
    return a.grade - b.grade;
  });

  // dailyRows: 後方互換のため列配置を維持（cycleOrder/cycleLabel は空値）
  const dailyRows = dailyEntries.map(function(entry) {
    return [
      entry.date,
      entry.fiscalYear,
      0,
      '',
      entry.weekKey,
      entry.grade,
      entry.plannedSessions,
      entry.elapsedFlag,
      entry.generatedAt
    ];
  });

  return {
    fiscalYear: normalizedFiscalYear,
    startDate: planStartDate || fiscalRange.startDate,
    endDate: planEndDate || fiscalRange.endDate,
    generatedAt: generatedAt,
    dailyPlanCount: dailyRows.length,
    dailyRows: dailyRows,
    planRows: planRows,
    totalsByGrade: totalsByGrade,
    reserveByGrade: reserveByGrade
  };
}

/**
 * 年度・学年別の学校日マップを構築
 * @param {number} fiscalYear - 対象年度
 * @param {Array<number>=} enabledWeekdays - 有効曜日配列（省略時はデフォルト）
 * @param {Date=} startDate - 実施期間の開始日（省略時は年度開始日）
 * @param {Date=} endDate - 実施期間の終了日（省略時は年度終了日）
 * @return {Object} 学年別日付配列
 */
function buildSchoolDayMapByGradeForFiscalYear(fiscalYear, enabledWeekdays, startDate, endDate) {
  const fiscalRange = getFiscalYearDateRange(fiscalYear);
  const rangeStart = startDate && startDate >= fiscalRange.startDate ? startDate : fiscalRange.startDate;
  const rangeEnd = endDate && endDate <= fiscalRange.endDate ? endDate : fiscalRange.endDate;
  const result = {};

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    result[grade] = [];
  }

  const rows = extractSchoolDayRows(rangeStart, rangeEnd, enabledWeekdays);
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
 * @param {Array<number>=} enabledWeekdays - 有効曜日配列（省略時はデフォルト月水金）
 * @return {Array<Object>} 日付・学年配列
 */
function extractSchoolDayRows(startDate, endDate, enabledWeekdays) {
  const sheet = getAnnualScheduleSheetOrThrow();
  const values = sheet.getDataRange().getValues();
  const rows = [];
  const allowedDays = Array.isArray(enabledWeekdays) && enabledWeekdays.length > 0
    ? enabledWeekdays
    : MODULE_DEFAULT_WEEKDAYS_ENABLED;
  const allowedSet = {};
  allowedDays.forEach(function(d) { allowedSet[d] = true; });

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);

    if (!date || date < startDate || date > endDate) {
      continue;
    }

    const day = date.getDay();
    if (!allowedSet[day]) {
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
 * セッションを学校週へ均等配分し、週内優先曜日で日付割当（1日1回上限）
 * @param {number} totalSessions - 年間総セッション
 * @param {Object} weekMap - 週キーごとの学校日配列
 * @return {Object} dateKey別セッション数
 */
function allocateSessionsToDateKeys(totalSessions, weekMap) {
  const allocations = {};
  const weekKeys = Object.keys(weekMap).sort();

  if (totalSessions <= 0 || weekKeys.length === 0) {
    return allocations;
  }

  // 週ごとの候補日（重複除去済み）と容量を確定
  const weekData = [];
  let totalCapacity = 0;
  weekKeys.forEach(function(weekKey) {
    const orderedDatesRaw = sortWeekDatesByPriority(weekMap[weekKey] || []);
    const uniqueMap = {};
    const orderedDates = [];

    orderedDatesRaw.forEach(function(date) {
      const dateKey = formatInputDate(date);
      if (uniqueMap[dateKey]) {
        return;
      }
      uniqueMap[dateKey] = true;
      orderedDates.push(date);
    });

    const capacity = orderedDates.length;
    totalCapacity += capacity;
    weekData.push({
      weekKey: weekKey,
      orderedDates: orderedDates,
      capacity: capacity,
      allocatedCount: 0
    });
  });

  if (totalCapacity <= 0) {
    return allocations;
  }

  // 1日1回上限により、割当可能数は実施可能日数総和まで
  const assignableSessions = Math.min(Math.round(totalSessions), totalCapacity);
  const basePerWeek = Math.floor(assignableSessions / weekData.length);
  const remainder = assignableSessions % weekData.length;

  // 第1段階: 既存ロジック同様に週へ均等配分し、週容量で上限適用
  weekData.forEach(function(week, index) {
    const extraCount = Math.floor((index + 1) * remainder / weekData.length) - Math.floor(index * remainder / weekData.length);
    const target = basePerWeek + extraCount;
    week.allocatedCount = Math.min(target, week.capacity);
  });

  // 第2段階: 祝日等で不足した分を、空きのある週へ再配分
  let allocatedTotal = 0;
  weekData.forEach(function(week) {
    allocatedTotal += week.allocatedCount;
  });

  let remaining = assignableSessions - allocatedTotal;
  while (remaining > 0) {
    let progressed = false;
    weekData.forEach(function(week) {
      if (remaining <= 0) {
        return;
      }
      if (week.allocatedCount < week.capacity) {
        week.allocatedCount += 1;
        remaining -= 1;
        progressed = true;
      }
    });
    if (!progressed) {
      break;
    }
  }

  // 週内優先曜日順の先頭日から割当（各日最大1）
  weekData.forEach(function(week) {
    for (let i = 0; i < week.allocatedCount; i++) {
      const dateKey = formatInputDate(week.orderedDates[i]);
      allocations[dateKey] = 1;
    }
  });

  if (totalSessions > assignableSessions) {
    Logger.log(
      '[INFO] 1日1回上限により割当上限に到達: requested=' + totalSessions + ', assigned=' + assignableSessions
    );
  }

  return allocations;
}

/**
 * 月別コマ目標からセッションを月ごとに配分
 * @param {Object} monthlyKoma - 月番号→コマ数マップ {4:3, 5:2, ...}
 * @param {Array<Date>} gradeDates - 学年の実施可能日リスト（ソート済み）
 * @return {Object} dateKey→セッション数マップ
 */
function allocateSessionsByMonth(monthlyKoma, gradeDates) {
  const allocations = {};

  // 日付を月別にグループ化
  const datesByMonth = {};
  gradeDates.forEach(function(date) {
    const month = date.getMonth() + 1;
    if (!datesByMonth[month]) {
      datesByMonth[month] = [];
    }
    datesByMonth[month].push(date);
  });

  // 月ごとに配分
  [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].forEach(function(month) {
    const koma = toNumberOrZero(monthlyKoma[month]);
    if (koma <= 0) {
      return;
    }
    const sessions = Math.max(0, Math.round(koma * 3));
    const monthDates = datesByMonth[month] || [];
    if (monthDates.length === 0) {
      Logger.log('[WARNING] ' + month + '月に実施可能日がありませんが、目標が' + koma + 'コマ設定されています');
      return;
    }

    const weekMap = buildWeekMapFromDates(monthDates);
    const monthAllocations = allocateSessionsToDateKeys(sessions, weekMap);

    Object.keys(monthAllocations).forEach(function(dateKey) {
      allocations[dateKey] = toNumberOrZero(allocations[dateKey]) + monthAllocations[dateKey];
    });
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

  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded();
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
    ensureDefaultAnnualTargetForFiscalYear(fiscalYear);
    const buildResult = buildDailyPlanFromAnnualTarget(fiscalYear, end);

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
  const controlSheet = initializeModuleHoursSheetsIfNeeded();
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const rows = readExceptionRows(controlSheet);

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
