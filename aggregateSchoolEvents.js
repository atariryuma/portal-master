/**
 * @fileoverview 学年別授業時数集計機能
 * @description 指定期間の学年別授業時数を低・中・高学年別に詳細集計します。
 *              日付選択ダイアログから期間を指定し、月別の詳細レポートを作成します。
 */

/**
 * 日付選択ダイアログを表示
 */
function aggregateSchoolEventsByGrade() {
  try {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createTemplateFromFile('DateSelector').evaluate();
    ui.showModalDialog(htmlOutput, '集計範囲の指定');
  } catch (error) {
    showAlert('ダイアログの表示に失敗しました: ' + error.toString(), 'エラー');
  }
}

/**
 * 低(1,2)、中(3,4)、高(5,6)の3シートに分けて集計
 * テンプレート名: '時数様式'
 */
function processAggregateSchoolEventsByGrade(startDate, endDate, gradeHours) {
  const dateRange = parseAndValidateAggregateDateRange(startDate, endDate);
  const startDateObj = dateRange.startDate;
  const endDateObj = dateRange.endDate;

  const gradeGroups = {
    '低学年': [1, 2],
    '中学年': [3, 4],
    '高学年': [5, 6],
  };

  const categories = EVENT_CATEGORIES;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = getAnnualScheduleSheet();
  if (!srcSheet) {
    throw new Error('年間行事予定表シートが見つからないか、データが不完全です。');
  }
  const data = srcSheet.getDataRange().getValues();

  const monthKeys = listMonthKeysInRange(startDateObj, endDateObj);
  let moduleCalculationError = '';

  let modulePlanMap = null;
  try {
    modulePlanMap = applyModuleExceptions(
      buildSchoolDayPlanMap(startDateObj, endDateObj),
      endDateObj
    );
  } catch (error) {
    moduleCalculationError = error.toString();
    Logger.log('[WARNING] MOD列の算出に失敗したため、R列の既存値を保持します: ' + moduleCalculationError);
  }

  const templateSheet = getSheetByNameOrThrow(JISUU_TEMPLATE.SHEET_NAME);
  templateSheet.hideSheet();

  Object.keys(gradeGroups).forEach(function(groupName) {
    const grades = gradeGroups[groupName];

    const newSheetName = groupName;
    let newSheet = ss.getSheetByName(newSheetName);
    let preservedModValuesByGrade = null;
    if (!modulePlanMap && newSheet) {
      preservedModValuesByGrade = captureExistingModValuesByMonth(
        newSheet,
        monthKeys,
        grades,
        JISUU_TEMPLATE.GRADE_BLOCK_HEIGHT,
        JISUU_TEMPLATE.MOD_COLUMN_INDEX
      );
    }
    if (!newSheet) {
      newSheet = templateSheet.copyTo(ss).setName(newSheetName);
    } else {
      newSheet.clear();
      templateSheet.getRange('A1:Z100').copyTo(newSheet.getRange('A1:Z100'));
    }
    newSheet.showSheet();

    // データ計算フェーズ（API呼び出しなし）
    const gradeWriteOps = grades.map(function(grade, idx) {
      const blockOffset = idx * JISUU_TEMPLATE.GRADE_BLOCK_HEIGHT;
      const results = collectMonthlyResultsForGrade_(data, grade, startDateObj, endDateObj, monthKeys, categories);
      const output = buildGradeOutputRows_(monthKeys, results, modulePlanMap, preservedModValuesByGrade, grade);

      return {
        gradeLabelRow: JISUU_TEMPLATE.GRADE_LABEL_ROW + blockOffset,
        gradeLabel: '【' + grade + '年】',
        standardHourRow: JISUU_TEMPLATE.STANDARD_HOUR_ROW + blockOffset,
        standardHour: gradeHours[grade],
        dataStartRow: JISUU_TEMPLATE.DATA_START_ROW + blockOffset,
        output: output
      };
    });

    // 書き込みフェーズ（API呼び出しを集約）
    gradeWriteOps.forEach(function(op) {
      newSheet.getRange(op.gradeLabelRow, 1).setValue(op.gradeLabel);
      newSheet.getRange(op.standardHourRow, 3).setValue(op.standardHour);

      if (op.output.batchData.length > 0) {
        newSheet.getRange(op.dataStartRow, 1, op.output.batchData.length, 14).setValues(op.output.batchData);

        const modRange = newSheet.getRange(op.dataStartRow, JISUU_TEMPLATE.MOD_COLUMN_INDEX, op.output.modValues.length, 1);
        modRange.setNumberFormat(JISUU_TEMPLATE.MOD_FRACTION_FORMAT);
        modRange.setValues(op.output.modValues);
      }
    });
  });

  if (moduleCalculationError) {
    showAlert(
      '時数様式の集計は完了しましたが、MOD列（R列）の算出に失敗したため既存値を保持しました。\n詳細: ' + moduleCalculationError,
      '警告'
    );
  }
}

/**
 * 指定学年のデータ行を走査し、月別の授業時数・カテゴリ別集計結果を返す
 * @param {Array<Array<*>>} data - シート全行データ
 * @param {number} grade - 対象学年（1-6）
 * @param {Date} startDateObj - 集計開始日
 * @param {Date} endDateObj - 集計終了日
 * @param {Array<string>} monthKeys - 対象月キー一覧（yyyy-MM）
 * @param {Object} categories - カテゴリ名→略称マップ
 * @return {Object} monthKey → カテゴリ別カウントのマップ
 */
function collectMonthlyResultsForGrade_(data, grade, startDateObj, endDateObj, monthKeys, categories) {
  const results = {};

  monthKeys.forEach(function(monthKey) {
    const entry = { "授業時数": 0 };
    Object.keys(categories).forEach(function(category) {
      entry[category] = 0;
    });
    entry["対象日数"] = 0;
    results[monthKey] = entry;
  });

  // カテゴリ略称→カテゴリ名の逆引きマップを事前構築（O(n*m)→O(n)に最適化）
  const abbreviationToCategory = buildAbbreviationToCategoryMap(categories);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);
    if (!date) continue;

    if (date >= startDateObj && date <= endDateObj) {
      const monthKey = formatMonthKey(date);
      if (Number(row[SCHEDULE_COLUMNS.GRADE]) === grade) {
        let hasClass = false;
        for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
          const cellValue = row[j];
          if (cellValue === "○") {
            results[monthKey]["授業時数"]++;
            hasClass = true;
          } else if (cellValue && Object.prototype.hasOwnProperty.call(abbreviationToCategory, cellValue)) {
            results[monthKey][abbreviationToCategory[cellValue]]++;
            hasClass = true;
          }
        }

        if (hasClass) {
          results[monthKey]["対象日数"]++;
        }
      }
    }
  }

  return results;
}

/**
 * 月別集計結果からシート出力用の2D配列とMOD値配列を構築
 * @param {Array<string>} monthKeys - 対象月キー一覧（yyyy-MM）
 * @param {Object} results - collectMonthlyResultsForGrade_ の戻り値
 * @param {Object|null} modulePlanMap - モジュール計画マップ（算出失敗時はnull）
 * @param {Object|null} preservedModValuesByGrade - 既存MOD値の退避データ
 * @param {number} grade - 対象学年（1-6）
 * @return {Object} { batchData: Array<Array<*>>, modValues: Array<Array<*>> }
 */
function buildGradeOutputRows_(monthKeys, results, modulePlanMap, preservedModValuesByGrade, grade) {
  const batchData = [];
  const modValues = [];

  monthKeys.forEach(function(monthKey) {
    if (results[monthKey]) {
      const rowData = [
        monthKey,                              // A列: 年月
        results[monthKey]["対象日数"],          // B列: 対象日数
        results[monthKey]["授業時数"],          // C列: 授業時数
        results[monthKey]["儀式"],              // D列: 儀式
        results[monthKey]["文化"],              // E列: 文化
        results[monthKey]["保健"],              // F列: 保健
        results[monthKey]["遠足"],              // G列: 遠足
        results[monthKey]["勤労"],              // H列: 勤労
        '',                                      // I列: 空欄
        results[monthKey]["欠時数"],            // J列: 欠時数
        results[monthKey]["児童会"],            // K列: 児童会
        results[monthKey]["クラブ"],            // L列: クラブ
        results[monthKey]["委員会活動"],        // M列: 委員会活動
        results[monthKey]["補習"]               // N列: 補習
      ];
      batchData.push(rowData);

      let modValue = '';
      if (modulePlanMap) {
        modValue = getModuleActualUnitsForMonth(modulePlanMap, monthKey, grade);
      } else if (preservedModValuesByGrade &&
          Object.prototype.hasOwnProperty.call(preservedModValuesByGrade, grade) &&
          Object.prototype.hasOwnProperty.call(preservedModValuesByGrade[grade], monthKey)) {
        modValue = preservedModValuesByGrade[grade][monthKey];
      }
      modValues.push([modValue]);
    }
  });

  return { batchData: batchData, modValues: modValues };
}

/**
 * 学年別集計の期間入力を検証してDateへ変換
 */
function parseAndValidateAggregateDateRange(startDate, endDate) {
  const startDateObj = normalizeToDate(startDate);
  const endDateObj = normalizeToDate(endDate);

  if (!startDateObj || !endDateObj) {
    throw new Error('入力された日付が無効です。');
  }
  if (startDateObj > endDateObj) {
    throw new Error('開始日は終了日以前の日付を指定してください。');
  }

  return {
    startDate: startDateObj,
    endDate: endDateObj
  };
}

/**
 * 既存シートのR列（MOD）を月キー・学年単位で退避
 */
function captureExistingModValuesByMonth(sheet, monthKeys, grades, blockHeight, modColumnIndex) {
  const valuesByGrade = {};
  if (!sheet || !Array.isArray(monthKeys) || monthKeys.length === 0 || !Array.isArray(grades)) {
    return valuesByGrade;
  }

  const blockRowCapacity = Number(blockHeight);
  const fallbackScanRowCount = Math.max(monthKeys.length, 24);
  const scanRowCount = Number.isFinite(blockRowCapacity) && blockRowCapacity > 0
    ? Math.min(fallbackScanRowCount, Math.floor(blockRowCapacity))
    : fallbackScanRowCount;

  grades.forEach(function(grade, index) {
    const rowStart = JISUU_TEMPLATE.DATA_START_ROW + (index * blockHeight);
    const monthValues = sheet.getRange(rowStart, 1, scanRowCount, 1).getValues();
    const modValues = sheet.getRange(rowStart, modColumnIndex, scanRowCount, 1).getValues();
    const existingByMonth = {};

    for (let i = 0; i < scanRowCount; i++) {
      const key = normalizeAggregateMonthKey(monthValues[i][0]);
      if (key) {
        existingByMonth[key] = modValues[i][0];
      }
    }

    const monthMap = {};
    monthKeys.forEach(function(monthKey) {
      if (Object.prototype.hasOwnProperty.call(existingByMonth, monthKey)) {
        monthMap[monthKey] = existingByMonth[monthKey];
      }
    });

    valuesByGrade[grade] = monthMap;
  });

  return valuesByGrade;
}

/**
 * 集計シートの月キーセル値を yyyy-MM に正規化
 */
function normalizeAggregateMonthKey(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }

  return String(value === null || value === undefined ? '' : value).trim();
}

/**
 * MOD月次マップから対象月・学年の実績値（45分コマ換算）を取得
 */
function getModuleActualUnitsForMonth(modulePlanMap, monthKey, grade) {
  if (!modulePlanMap || !modulePlanMap.byMonth || !modulePlanMap.byMonth[monthKey]) {
    return 0;
  }

  const entry = modulePlanMap.byMonth[monthKey][grade];
  if (!entry) {
    return 0;
  }

  const value = Number(entry.actual_units);
  return Number.isFinite(value) ? value : 0;
}
