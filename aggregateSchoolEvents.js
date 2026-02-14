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
  const templateSheetName = '時数様式';
  const GRADE_BLOCK_HEIGHT = 21; // 時数様式シート内の学年ブロック間の行数
  const MOD_COLUMN_INDEX = 18; // R列

  // 学年グループ化: 低(1,2)、中(3,4)、高(5,6)
  const gradeGroups = {
    '低学年': [1, 2],
    '中学年': [3, 4],
    '高学年': [5, 6],
  };

  // 行事カテゴリ(従来通り)
  // 共通関数からカテゴリーを取得
  const categories = EVENT_CATEGORIES;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = getAnnualScheduleSheet(); // 共通関数を使用
  if (!srcSheet) {
    showAlert('年間行事予定表シートが見つからないか、データが不完全です。');
    return;
  }
  const data = srcSheet.getDataRange().getValues();

  const startDateObj = new Date(startDate);
  const endDateObj = new Date(endDate);

  if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
    showAlert('入力された日付が無効です。');
    return;
  }
  if (startDateObj > endDateObj) {
    showAlert('開始日は終了日以前の日付を指定してください。');
    return;
  }

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

  const templateSheet = getSheetByNameOrThrow(templateSheetName);
  // テンプレートは表示しない
  templateSheet.hideSheet();

  /**
   * gradeGroups でループ: 低学年、中学年、高学年の3パターン
   */
  Object.keys(gradeGroups).forEach(function(groupName) {
    // 例: groupName='低学年'、grades=[1,2]
    const grades = gradeGroups[groupName];

    // 新しいシート名: 例) '低学年'
    const newSheetName = groupName;
    let newSheet = ss.getSheetByName(newSheetName);
    let preservedModValuesByGrade = null;
    if (!modulePlanMap && newSheet) {
      preservedModValuesByGrade = captureExistingModValuesByMonth(
        newSheet,
        monthKeys,
        grades,
        GRADE_BLOCK_HEIGHT,
        MOD_COLUMN_INDEX
      );
    }
    if (!newSheet) {
      newSheet = templateSheet.copyTo(ss).setName(newSheetName);
    } else {
      // 既存があればクリアしてテンプレートをコピー
      newSheet.clear();
      templateSheet.getRange('A1:Z100').copyTo(newSheet.getRange('A1:Z100'));
    }
    newSheet.showSheet();

    // 今回は「grades.length=2」想定 (低学年なら1,2年)
    // 1つ目の学年をシートの上部ブロック (A2, row=4～)
    // 2つ目の学年をシートの下部ブロック (A23, row=25～)
    grades.forEach(function(grade, idx) {
      // idx=0 → 1学年目, idx=1 → 2学年目
      // ブロックの行オフセット(上から何行ずらすか)
      // 例: idx=0 → 0行、idx=1 → +21行
      const blockOffset = idx * GRADE_BLOCK_HEIGHT;

      // 学年ラベルの配置: A2（1学年目）、A23（2学年目）
      const gradeCellRow = 2 + blockOffset;
      newSheet.getRange(gradeCellRow, 1).setValue('【' + grade + '年】');

      // 標準授業時数の配置: C17（1学年目）、C38（2学年目）
      const standardHourRow = 17 + blockOffset;
      newSheet.getRange(standardHourRow, 3).setValue(gradeHours[grade]);

      // 集計結果を格納するオブジェクト results
      const results = {};

      // startDateObj から endDateObj まで、1ヶ月ずつ進める
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

      // データ集計
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const date = new Date(row[SCHEDULE_COLUMNS.DATE]);
        if (isNaN(date.getTime())) continue;

        if (date >= startDateObj && date <= endDateObj) {
          const monthKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
          // 学年列から学年を取得: grade(1～6)
          if (row[SCHEDULE_COLUMNS.GRADE] == grade) {
            let hasClass = false;
            // 授業(○)をカウント
            for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
              if (row[j] == "○") {
                results[monthKey]["授業時数"]++;
                hasClass = true;
              }
            }
            // 行事カテゴリをカウント
            Object.keys(categories).forEach(function(category) {
              for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
                if (row[j] == categories[category]) {
                  results[monthKey][category]++;
                  hasClass = true;
                }
              }
            });

            // その日になにかしら(授業 or 行事)があったら 対象日数+1
            if (hasClass) {
              results[monthKey]["対象日数"]++;
            }
          }
        }
      }

      // シートへの書き込み
      // 例: 上ブロックは row=4, 下ブロックは row=25 => row=4+blockOffset
      // R列（18列目）には MOD の月次実績（45分コマ換算）を出力
      let rowIndexBase = 4 + blockOffset;
      monthKeys.forEach(function(monthKey2) {
        if (results[monthKey2]) {
          newSheet.getRange(rowIndexBase, 1).setValue(monthKey2);    // A列: 年月
          newSheet.getRange(rowIndexBase, 2).setValue(results[monthKey2]["対象日数"]);  // B列: 対象日数
          newSheet.getRange(rowIndexBase, 3).setValue(results[monthKey2]["授業時数"]);  // C列: 授業時数
          newSheet.getRange(rowIndexBase, 4).setValue(results[monthKey2]["儀式"]);      // D列: 儀式
          newSheet.getRange(rowIndexBase, 5).setValue(results[monthKey2]["文化"]);      // E列: 文化
          newSheet.getRange(rowIndexBase, 6).setValue(results[monthKey2]["保健"]);      // F列: 保健
          newSheet.getRange(rowIndexBase, 7).setValue(results[monthKey2]["遠足"]);      // G列: 遠足
          newSheet.getRange(rowIndexBase, 8).setValue(results[monthKey2]["勤労"]);      // H列: 勤労
          // I列（9列目）: 空欄
          newSheet.getRange(rowIndexBase, 10).setValue(results[monthKey2]["欠時数"]);   // J列: 欠時数
          newSheet.getRange(rowIndexBase, 11).setValue(results[monthKey2]["児童会"]);   // K列: 児童会
          newSheet.getRange(rowIndexBase, 12).setValue(results[monthKey2]["クラブ"]);   // L列: クラブ
          newSheet.getRange(rowIndexBase, 13).setValue(results[monthKey2]["委員会活動"]); // M列: 委員会活動
          newSheet.getRange(rowIndexBase, 14).setValue(results[monthKey2]["補習"]);     // N列: 補習
          // O列～Q列（15～17列目）: 空欄
          if (modulePlanMap) {
            newSheet.getRange(rowIndexBase, MOD_COLUMN_INDEX).setValue(
              getModuleActualUnitsForMonth(modulePlanMap, monthKey2, grade)
            ); // R列: MOD
          } else if (preservedModValuesByGrade &&
              Object.prototype.hasOwnProperty.call(preservedModValuesByGrade, grade) &&
              Object.prototype.hasOwnProperty.call(preservedModValuesByGrade[grade], monthKey2)) {
            newSheet.getRange(rowIndexBase, MOD_COLUMN_INDEX).setValue(
              preservedModValuesByGrade[grade][monthKey2]
            ); // R列: MOD（既存値保持）
          }

          rowIndexBase++;
        }
      });
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
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array<string>} 月キー配列
 */
function buildMonthKeysForAggregate(startDate, endDate) {
  const keys = [];
  let cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const lastMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  while (cursor <= lastMonth) {
    keys.push(Utilities.formatDate(cursor, 'Asia/Tokyo', 'yyyy-MM'));
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }

  return keys;
}

/**
 * 既存シートのR列（MOD）を月キー・学年単位で退避
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<string>} monthKeys - 対象月キー配列
 * @param {Array<number>} grades - 対象学年配列
 * @param {number} blockHeight - 学年ブロック高さ
 * @param {number} modColumnIndex - R列インデックス
 * @return {Object} 学年ごとの月キー値マップ
 */
function captureExistingModValuesByMonth(sheet, monthKeys, grades, blockHeight, modColumnIndex) {
  const valuesByGrade = {};
  if (!sheet || !Array.isArray(monthKeys) || monthKeys.length === 0 || !Array.isArray(grades)) {
    return valuesByGrade;
  }

  grades.forEach(function(grade, index) {
    const rowStart = 4 + (index * blockHeight);
    const scanRowCount = Math.max(monthKeys.length, 24);
    const monthValues = sheet.getRange(rowStart, 1, scanRowCount, 1).getValues();
    const modValues = sheet.getRange(rowStart, modColumnIndex, scanRowCount, 1).getValues();
    const existingByMonth = {};

    for (let i = 0; i < scanRowCount; i++) {
      const key = String(monthValues[i][0] || '').trim();
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
 * MOD月次マップから対象月・学年の実績値（45分コマ換算）を取得
 * @param {Object} modulePlanMap - buildSchoolDayPlanMap/applyModuleExceptionsの結果
 * @param {string} monthKey - yyyy-MM
 * @param {number} grade - 学年
 * @return {number} MOD実績
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
