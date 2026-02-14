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

const CUMULATIVE_EVENT_CATEGORIES = ["儀式", "文化", "保健", "遠足", "勤労", "欠時数", "児童会", "クラブ", "委員会活動"];

// 結果をシートに書き込む関数
function writeResultsToSheet(sheet, grade, results) {
  const gradeRow = 3 + (grade - 1); // 各学年の行番号（C3:L8の範囲に対応）

  Object.keys(results).forEach(function(key) {
    const column = CUMULATIVE_COLUMN_MAP[key];
    if (!column) {
      Logger.log('[WARNING] 累計時数シートに未定義のカテゴリをスキップしました: ' + key);
      return;
    }
    sheet.getRange(gradeRow, column).setValue(results[key]);
  });
}

// メイン関数
function calculateCumulativeHours() {
  try {
    const grades = [1, 2, 3, 4, 5, 6];
    const categories = getCumulativeCategoryMap();

    const srcSheet = getAnnualScheduleSheetOrThrow(); // 共通関数を使用してエラーハンドリング

    const data = srcSheet.getDataRange().getValues();
    const cumulativeSheet = getSheetByNameOrThrow('累計時数');
    const thisSaturday = getCurrentOrNextSaturday(); // 共通関数を使用
    const formattedDate = formattedDateMessage(thisSaturday);

    // 累計日付をシートに設定
    cumulativeSheet.getRange('A1').setValue(formattedDate);
    Logger.log("この週の土曜日の日付: " + Utilities.formatDate(thisSaturday, Session.getScriptTimeZone(), 'yyyy/MM/dd'));

    grades.forEach(function(grade) {
      const results = calculateResultsForGrade(data, grade, thisSaturday, categories);
      writeResultsToSheet(cumulativeSheet, grade, results);
    });

    // 既存の累計計算にモジュール学習計画を統合
    syncModuleHoursWithCumulative(thisSaturday);

    // アラートを表示（共通関数を使用）
    showAlert(formattedDate + 'を計算しました。モジュール学習計画も更新済みです。');

  } catch (error) {
    showAlert(error.message, 'エラー');
  }
}

// getSheetByNameOrThrow関数はcommon.jsに移行済み
// getNextSaturday関数はcommon.jsのgetCurrentOrNextSaturdayに移行済み

// 日付のメッセージをフォーマット
function formattedDateMessage(date) {
  return (date.getMonth() + 1) + "月" + date.getDate() + "日までの累計時数";
}

// showAlert関数はcommon.jsに移行済み

// 指定学年の累計結果を計算
function calculateResultsForGrade(data, grade, endDate, categories) {
  const results = {};
  Object.keys(CUMULATIVE_COLUMN_MAP).forEach(function(key) {
    results[key] = 0;
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateValue = row[SCHEDULE_COLUMNS.DATE];
    if (!dateValue) continue; // 日付がない場合はスキップ

    const date = new Date(dateValue);

    // 日付が無効または集計対象日付を超える場合はスキップ
    if (isNaN(date.getTime()) || date > endDate) continue;

    // 指定学年の行のみを対象にカウント
    if (row[SCHEDULE_COLUMNS.GRADE] == grade) {
      // "○" のカウントを正しく処理
      for (let j = SCHEDULE_COLUMNS.DATA_START; j <= SCHEDULE_COLUMNS.DATA_END; j++) {
        if (row[j] == "○") results["授業時数"]++;
      }

      // カテゴリーごとのカウント（累計時数シートに出力するカテゴリのみ）
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
