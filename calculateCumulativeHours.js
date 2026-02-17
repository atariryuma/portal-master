/**
 * @fileoverview 累計時数計算機能
 * @description 各学年の授業時数を累計計算し、累計時数シートに出力します。
 *              直近の土曜日までの累計を計算し、共通関数による改善された
 *              土曜日計算ロジックを使用します。
 */

// カラム定数はcommon.jsのSCHEDULE_COLUMNSを使用
const CUMULATIVE_COLUMN_MAP = Object.freeze({
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
});

// CUMULATIVE_EVENT_CATEGORIESはcommon.jsで定義（GASファイル読み込み順序の制約）

function writeResultsToSheet(sheet, grade, results) {
  const gradeRow = CUMULATIVE_SHEET.GRADE_START_ROW + (grade - 1);

  // バッチ書き込み: 全カテゴリの値を配列に構築して一括設定
  const sortedKeys = Object.keys(CUMULATIVE_COLUMN_MAP).sort(function(a, b) {
    return CUMULATIVE_COLUMN_MAP[a] - CUMULATIVE_COLUMN_MAP[b];
  });
  const minCol = CUMULATIVE_COLUMN_MAP[sortedKeys[0]];
  const maxCol = CUMULATIVE_COLUMN_MAP[sortedKeys[sortedKeys.length - 1]];
  const row = new Array(maxCol - minCol + 1).fill('');

  sortedKeys.forEach(function(key) {
    row[CUMULATIVE_COLUMN_MAP[key] - minCol] = results[key] !== undefined ? results[key] : 0;
  });

  sheet.getRange(gradeRow, minCol, 1, row.length).setValues([row]);
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

    grades.forEach(function(grade) {
      const results = calculateResultsForGrade(data, grade, thisSaturday, categories);
      writeResultsToSheet(cumulativeSheet, grade, results);
    });

    syncModuleHoursWithCumulative(thisSaturday);

    showAlert(formattedDate + 'を計算しました。モジュール学習計画も更新済みです。', '通知');

  } catch (error) {
    showAlert(error.message, 'エラー');
  }
}

/**
 * 年間行事予定表データから指定学年の累計時数を計算
 *
 * 集計ロジック:
 * - 「○」セルは通常授業の1時間としてカウント
 * - カテゴリ略称（「儀式」「文化」等）のセルは特別活動としてカウント
 * - 逆引きマップを事前構築し、データループをO(n)で処理する
 *   （マップなしだとO(n*m)の二重ループが必要）
 */
function calculateResultsForGrade(data, grade, endDate, categories) {
  const results = {};
  Object.keys(CUMULATIVE_COLUMN_MAP).forEach(function(key) {
    results[key] = 0;
  });

  // カテゴリ略称→カテゴリ名の逆引きマップを事前構築（ループ内検索を排除）
  const abbreviationToCategory = buildAbbreviationToCategoryMap(categories);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = normalizeToDate(row[SCHEDULE_COLUMNS.DATE]);
    if (!date || date > endDate) continue;

    if (Number(row[SCHEDULE_COLUMNS.GRADE]) === grade) {
      // 1回のカラムスキャンで授業時数とカテゴリを同時にカウント
      for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
        const cellValue = row[j];
        if (cellValue === "○") {
          results["授業時数"]++;
        } else if (cellValue && Object.prototype.hasOwnProperty.call(abbreviationToCategory, cellValue)) {
          results[abbreviationToCategory[cellValue]]++;
        }
      }
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
