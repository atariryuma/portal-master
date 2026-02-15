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
    const htmlOutput = HtmlService.createHtmlOutputFromFile('DateSelector');
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

  const monthKeys = buildMonthKeysForAggregate(startDateObj, endDateObj);
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

    grades.forEach(function(grade, idx) {
      const blockOffset = idx * JISUU_TEMPLATE.GRADE_BLOCK_HEIGHT;

      const gradeCellRow = JISUU_TEMPLATE.GRADE_LABEL_ROW + blockOffset;
      newSheet.getRange(gradeCellRow, 1).setValue('【' + grade + '年】');

      const standardHourRow = JISUU_TEMPLATE.STANDARD_HOUR_ROW + blockOffset;
      newSheet.getRange(standardHourRow, 3).setValue(gradeHours[grade]);

      const results = {};

      monthKeys.forEach(function(monthKey) {
        results[monthKey] = {
          "授業時数": 0,
          "儀式": 0,
          "文化": 0,
          "保健": 0,
          "遠足": 0,
          "勤労": 0,
          "欠時数": 0,
          "児童会": 0,
          "クラブ": 0,
          "委員会活動": 0,
          "補習": 0,
          "対象日数": 0
        };
      });

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);
        if (!date) continue;

        if (date >= startDateObj && date <= endDateObj) {
          const monthKey = formatMonthKey(date);
          if (Number(row[SCHEDULE_COLUMNS.GRADE]) === grade) {
            let hasClass = false;
            for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
              if (row[j] === "○") {
                results[monthKey]["授業時数"]++;
                hasClass = true;
              }
            }
            Object.keys(categories).forEach(function(category) {
              for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
                if (row[j] === categories[category]) {
                  results[monthKey][category]++;
                  hasClass = true;
                }
              }
            });

            if (hasClass) {
              results[monthKey]["対象日数"]++;
            }
          }
        }
      }

      // バッチ書き込み: 月別データを2D配列として構築し一括書き込み
      const rowIndexBase = JISUU_TEMPLATE.DATA_START_ROW + blockOffset;
      const batchData = [];
      const modValues = [];

      monthKeys.forEach(function(monthKey2) {
        if (results[monthKey2]) {
          const rowData = [
            monthKey2,                              // A列: 年月
            results[monthKey2]["対象日数"],          // B列: 対象日数
            results[monthKey2]["授業時数"],          // C列: 授業時数
            results[monthKey2]["儀式"],              // D列: 儀式
            results[monthKey2]["文化"],              // E列: 文化
            results[monthKey2]["保健"],              // F列: 保健
            results[monthKey2]["遠足"],              // G列: 遠足
            results[monthKey2]["勤労"],              // H列: 勤労
            '',                                      // I列: 空欄
            results[monthKey2]["欠時数"],            // J列: 欠時数
            results[monthKey2]["児童会"],            // K列: 児童会
            results[monthKey2]["クラブ"],            // L列: クラブ
            results[monthKey2]["委員会活動"],        // M列: 委員会活動
            results[monthKey2]["補習"]               // N列: 補習
          ];
          batchData.push(rowData);

          // MOD値を別途計算
          let modValue = '';
          if (modulePlanMap) {
            modValue = getModuleActualUnitsForMonth(modulePlanMap, monthKey2, grade);
          } else if (preservedModValuesByGrade &&
              Object.prototype.hasOwnProperty.call(preservedModValuesByGrade, grade) &&
              Object.prototype.hasOwnProperty.call(preservedModValuesByGrade[grade], monthKey2)) {
            modValue = preservedModValuesByGrade[grade][monthKey2];
          }
          modValues.push([modValue]);
        }
      });

      if (batchData.length > 0) {
        // A-N列を一括書き込み (14列)
        newSheet.getRange(rowIndexBase, 1, batchData.length, 14).setValues(batchData);

        // R列(MOD)を一括書き込み
        const modRange = newSheet.getRange(rowIndexBase, JISUU_TEMPLATE.MOD_COLUMN_INDEX, modValues.length, 1);
        modRange.setNumberFormat(JISUU_TEMPLATE.MOD_FRACTION_FORMAT);
        modRange.setValues(modValues);
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
 * 集計対象期間の月キー一覧（yyyy-MM）を作成
 * listMonthKeysInRange (moduleHoursPlanning.js) へ委譲
 */
function buildMonthKeysForAggregate(startDate, endDate) {
  return listMonthKeysInRange(startDate, endDate);
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

  return String(value == null ? '' : value).trim();
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
