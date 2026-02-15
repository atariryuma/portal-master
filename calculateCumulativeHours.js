/**
 * @fileoverview 累計時数計算機能
 * @description 各学年の授業時数を累計計算し、累計時数シートに出力します。
 *              直近の土曜日までの累計を計算し、共通関数による改善された
 *              土曜日計算ロジックを使用します。
 */

// カラム定数はcommon.jsのSCHEDULE_COLUMNSを使用
const CUMULATIVE_COLUMN_MAP = {
  "授業時数": 3,
  "儀式": 4,
  "文化": 5,
  "保健": 6,
  "遠足": 7,
  "勤労": 8,
  "欠時数": 9,
  "児童会": 10,
  "クラブ": 11,
  "委員会活動": 12
};

// CUMULATIVE_EVENT_CATEGORIESはcommon.jsで定義（GASファイル読み込み順序の制約）

function writeResultsToSheet(sheet, grade, results) {
  const gradeRow = CUMULATIVE_SHEET.GRADE_START_ROW + (grade - 1);

  Object.keys(results).forEach(function(key) {
    const column = CUMULATIVE_COLUMN_MAP[key];
    if (!column) {
      Logger.log('[WARNING] 累計時数シートに未定義のカテゴリをスキップしました: ' + key);
      return;
    }
    sheet.getRange(gradeRow, column).setValue(results[key]);
  });
}

function calculateCumulativeHours() {
  try {
    const grades = [1, 2, 3, 4, 5, 6];
    const categories = getCumulativeCategoryMap();

    const srcSheet = getAnnualScheduleSheetOrThrow();

    const data = srcSheet.getDataRange().getValues();
    const cumulativeSheet = getSheetByNameOrThrow(CUMULATIVE_SHEET.NAME);
    const thisSaturday = getCurrentOrNextSaturday();
    const formattedDate = formatDateToJapanese(thisSaturday) + 'までの累計時数';

    cumulativeSheet.getRange(CUMULATIVE_SHEET.DATE_CELL).setValue(formattedDate);
    Logger.log("[DEBUG] この週の土曜日の日付: " + Utilities.formatDate(thisSaturday, Session.getScriptTimeZone(), 'yyyy/MM/dd'));

    grades.forEach(function(grade) {
      const results = calculateResultsForGrade(data, grade, thisSaturday, categories);
      writeResultsToSheet(cumulativeSheet, grade, results);
    });

    syncModuleHoursWithCumulative(thisSaturday);

    showAlert(formattedDate + 'を計算しました。モジュール学習計画も更新済みです。');

  } catch (error) {
    showAlert(error.message, 'エラー');
  }
}

function calculateResultsForGrade(data, grade, endDate, categories) {
  const results = {};
  Object.keys(CUMULATIVE_COLUMN_MAP).forEach(function(key) {
    results[key] = 0;
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateValue = row[SCHEDULE_COLUMNS.DATE];
    if (!dateValue) continue;

    const date = new Date(dateValue);

    if (isNaN(date.getTime()) || date > endDate) continue;

    if (row[SCHEDULE_COLUMNS.GRADE] == grade) {
      for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
        if (row[j] == "○") results["授業時数"]++;
      }

      Object.keys(categories).forEach(function (category) {
        for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
          if (row[j] == categories[category]) {
            results[category]++;
          }
        }
      });
    }
  }

  return results;
}

function getCumulativeCategoryMap() {
  const map = {};
  CUMULATIVE_EVENT_CATEGORIES.forEach(function(category) {
    if (Object.prototype.hasOwnProperty.call(EVENT_CATEGORIES, category)) {
      map[category] = EVENT_CATEGORIES[category];
    }
  });
  return map;
}
