/**
 * @fileoverview モジュール学習管理 - 配分アルゴリズム
 * @description クール計画から日次配分の構築、学校日マップ、例外集計を担当します。
 */

/** @type {?Object} 曜日優先度マップのキャッシュ */
let activeWeekdayPriorityMap_ = null;

/** @type {?Object} 学校日マップのキャッシュ（実行単位） */
let schoolDayMapCache_ = null;

/**
 * 保存された曜日優先度マップを読み込み（キャッシュ付き）
 * @return {Object} dayOfWeek → priorityIndex のマップ
 */
function loadWeekdayPriorityMap_() {
  if (activeWeekdayPriorityMap_) {
    return activeWeekdayPriorityMap_;
  }
  const settings = readModuleSettingsMap();
  const raw = settings[MODULE_SETTING_KEYS.WEEKDAY_PRIORITY];
  if (raw) {
    try {
      const days = JSON.parse(raw);
      if (Array.isArray(days) && days.length > 0) {
        const map = {};
        days.forEach(function(day, index) {
          map[Number(day)] = index;
        });
        activeWeekdayPriorityMap_ = map;
        return map;
      }
    } catch (e) {
      Logger.log('[WARNING] 曜日優先度設定の解析に失敗: ' + e.toString());
    }
  }
  activeWeekdayPriorityMap_ = MODULE_WEEKDAY_PRIORITY;
  return MODULE_WEEKDAY_PRIORITY;
}

/**
 * 曜日優先度キャッシュをリセット
 */
function resetWeekdayPriorityCache_() {
  activeWeekdayPriorityMap_ = null;
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

  upsertModuleSettingsValues({
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
 * @param {Array<Array<*>>=} existingRowsForFiscalYear - 事前取得済みの対象年度行
 * @return {boolean} 作成した場合true
 */
function ensureDefaultCyclePlanForFiscalYear(fiscalYear, controlSheet, existingRowsForFiscalYear) {
  const sheet = controlSheet || initializeModuleHoursSheetsIfNeeded().controlSheet;
  const existingRows = Array.isArray(existingRowsForFiscalYear)
    ? existingRowsForFiscalYear
    : readCyclePlanRowsByFiscalYear(sheet, fiscalYear);

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

  return toCyclePlansFromRows(fiscalYear, rows);
}

/**
 * 指定年度の計画行を計画オブジェクトへ変換
 * @param {number} fiscalYear - 対象年度
 * @param {Array<Array<*>>} rows - 計画行
 * @return {Array<Object>} 正規化済み計画
 */
function toCyclePlansFromRows(fiscalYear, rows) {
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
 * クール計画から日次計画を構築（保存はしない）
 * @param {number} fiscalYear - 対象年度
 * @param {Date|string} baseDate - 集計基準日
 * @return {Object} 構築結果
 */
function buildDailyPlanFromCyclePlan(fiscalYear, baseDate) {
  return buildDailyPlanFromCyclePlanInternal(fiscalYear, baseDate);
}

/**
 * クール計画から日次計画を構築（内部実装）
 * @param {number} fiscalYear - 対象年度
 * @param {Date|string} baseDate - 集計基準日
 * @param {?Object} options - 実行オプション
 * @return {Object} 構築結果
 */
function buildDailyPlanFromCyclePlanInternal(fiscalYear, baseDate, options) {
  const normalizedFiscalYear = Number(fiscalYear);
  const cutoffDate = normalizeToDate(baseDate) || normalizeToDate(new Date());
  const generatedAt = new Date();
  const fiscalRange = getFiscalYearDateRange(normalizedFiscalYear);
  const weekStart = getWeekStartMonday(cutoffDate);
  const controlSheet = options && options.controlSheet ? options.controlSheet : null;
  const cyclePlans = (options && options.cyclePlans)
    ? options.cyclePlans
    : loadCyclePlanForFiscalYear(normalizedFiscalYear, controlSheet);
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

  const carryOverByGrade = {};
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    carryOverByGrade[grade] = 0;
  }

  cyclePlans.forEach(function(plan) {
    const cycleLabel = plan.startMonth + '-' + plan.endMonth;
    const cycleMonthSet = buildCycleMonthKeySetForFiscalYear(normalizedFiscalYear, plan.startMonth, plan.endMonth);

    for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
      const plannedKoma = toNumberOrZero(plan.gradeKoma[grade]);
      const plannedSessions = Math.max(0, Math.round(plannedKoma * 3));
      const effectiveSessions = plannedSessions + carryOverByGrade[grade];
      const gradeDates = schoolDayMap[grade].filter(function(date) {
        return !!cycleMonthSet[formatMonthKey(date)];
      });

      const allocResult = allocateSessionsToDateKeys(effectiveSessions, gradeDates);
      const allocations = allocResult.allocations;
      carryOverByGrade[grade] = allocResult.overflow;
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

      if (effectiveSessions > 0 && allocatedDateKeys.length === 0) {
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

  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    if (carryOverByGrade[grade] > 0) {
      Logger.log('[WARNING] 配分しきれないセッション: FY' + normalizedFiscalYear + ', grade=' + grade + ', overflow=' + carryOverByGrade[grade]);
    }
  }

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

  return {
    fiscalYear: normalizedFiscalYear,
    startDate: fiscalRange.startDate,
    endDate: fiscalRange.endDate,
    generatedAt: generatedAt,
    dailyPlanCount: dailyRows.length,
    dailyRows: dailyRows,
    planRows: planRows,
    totalsByGrade: totalsByGrade,
    overflowByGrade: carryOverByGrade
  };
}

/**
 * 年度・学年別の学校日マップを構築（実行単位キャッシュ付き）
 * @param {number} fiscalYear - 対象年度
 * @return {Object} 学年別日付配列
 */
function buildSchoolDayMapByGradeForFiscalYear(fiscalYear) {
  const cacheKey = String(fiscalYear);
  if (schoolDayMapCache_ && schoolDayMapCache_.fiscalYear === cacheKey) {
    return schoolDayMapCache_.data;
  }

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

  schoolDayMapCache_ = { fiscalYear: cacheKey, data: result };
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
 * セッションを曜日優先度ごとに3等分し、各曜日内でBresenham均等配分（1日1セッション上限）
 * 高優先曜日から順に充填し、余りは優先度順に+1ずつ配分する。
 * 特定曜日の日数が不足した場合、溢れ分は次の優先曜日へ繰り越す。
 * @param {number} totalSessions - クール総セッション
 * @param {Array<Date>} dates - クール内の学校日（優先曜日以外も含む）
 * @return {Object} {allocations: dateKey別セッション数, overflow: 配分しきれなかったセッション数}
 */
function allocateSessionsToDateKeys(totalSessions, dates) {
  const allocations = {};

  if (totalSessions <= 0 || dates.length === 0) {
    return { allocations: allocations, overflow: Math.max(0, totalSessions) };
  }

  const priorityGroups = {};
  dates.forEach(function(date) {
    const pri = weekdayPriority(date.getDay());
    if (pri === 99) { return; }
    if (!priorityGroups[pri]) {
      priorityGroups[pri] = [];
    }
    priorityGroups[pri].push(date);
  });

  const priorities = Object.keys(priorityGroups).map(Number).sort(function(a, b) { return a - b; });

  if (priorities.length === 0) {
    return { allocations: allocations, overflow: totalSessions };
  }

  const basePerPriority = Math.floor(totalSessions / priorities.length);
  let extraSessions = totalSessions % priorities.length;
  let carry = 0;

  priorities.forEach(function(pri) {
    const sortedDates = priorityGroups[pri].sort(function(a, b) {
      return a.getTime() - b.getTime();
    });

    let target = basePerPriority + carry;
    if (extraSessions > 0) {
      target += 1;
      extraSessions -= 1;
    }
    carry = 0;

    if (target <= 0) { return; }

    const available = sortedDates.length;

    if (target >= available) {
      sortedDates.forEach(function(date) {
        allocations[formatInputDate(date)] = 1;
      });
      carry = target - available;
    } else {
      distributeByBresenham(sortedDates, target, allocations);
    }
  });

  return { allocations: allocations, overflow: carry };
}

/**
 * Bresenhamアルゴリズムで日付配列から均等にtarget個を選択して配分
 * @param {Array<Date>} sortedDates - 時系列ソート済み日付配列
 * @param {number} target - 配分するセッション数
 * @param {Object} allocations - 配分先オブジェクト（dateKey → セッション数）
 */
function distributeByBresenham(sortedDates, target, allocations) {
  const total = sortedDates.length;
  for (let i = 0; i < total; i++) {
    if (Math.floor((i + 1) * target / total) > Math.floor(i * target / total)) {
      allocations[formatInputDate(sortedDates[i])] = 1;
    }
  }
}

/**
 * 曜日優先度を返却（設定値またはデフォルト: 月→水→金）
 * 優先度マップに含まれない曜日は 99 を返す（配分対象外）。
 * @param {number} dayOfWeek - Date#getDay()
 * @return {number} 優先度インデックス（小さいほど優先、99=対象外）
 */
function weekdayPriority(dayOfWeek) {
  const map = loadWeekdayPriorityMap_();
  if (Object.prototype.hasOwnProperty.call(map, dayOfWeek)) {
    return map[dayOfWeek];
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
    const buildResult = buildDailyPlanFromCyclePlanInternal(fiscalYear, end);

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
