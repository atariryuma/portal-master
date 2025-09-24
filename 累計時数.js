// 定数の定義
const DATE_COLUMN = 0;           // 日付列（A列）
const GRADE_COLUMN = 19;         // 学年列（T列）
const DATA_COLUMNS_START = 20;   // データ開始列（U列）
const DATA_COLUMNS_END = 25;     // データ終了列（Z列）

// 結果をシートに書き込む関数
function writeResultsToSheet(sheet, grade, results) {
  const gradeRow = 3 + (grade - 1); // 各学年の行番号（C3:L8の範囲に対応）
  const columnMap = {
    "授業時数": 3, "儀式": 4, "文化": 5, "保健": 6, "遠足": 7,
    "勤労": 8, "欠時数": 9, "児童会": 10, "クラブ": 11, "委員会活動": 12
  };

  Object.keys(results).forEach(function(key) {
    sheet.getRange(gradeRow, columnMap[key]).setValue(results[key]);
  });
}

// メイン関数
function calculateCumulativeHours() {
  try {
    const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const grades = [1, 2, 3, 4, 5, 6];
    // 共通関数からカテゴリーを取得し、授業時数を追加
    const categories = Object.assign({"授業時数": "授業時数"}, EVENT_CATEGORIES);

    const srcSheet = getAnnualScheduleSheetOrThrow(); // 共通関数を使用してエラーハンドリング

    const data = srcSheet.getDataRange().getValues();
    const cumulativeSheet = getSheetByNameOrThrow('累計時数');
    const thisSaturday = getCurrentOrNextSaturday(); // 共通関数を使用
    const formattedDate = formattedDateMessage(thisSaturday);

    // 累計日付をシートに設定
    cumulativeSheet.getRange('A1').setValue(formattedDate);
    console.log("この週の土曜日の日付: " + thisSaturday.toLocaleDateString());

    // アラートを表示（共通関数を使用）
    showAlert(formattedDate + 'を計算しました。自動計算は引き続き行います。');

    grades.forEach(function(grade) {
      const results = calculateResultsForGrade(data, grade, thisSaturday, categories);
      writeResultsToSheet(cumulativeSheet, grade, results);
    });

  } catch (error) {
    showAlert(error.message, 'エラー');
  }
}

// シートの取得とエラーチェック
function getSheetByNameOrThrow(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + 'シートが見つかりません。');
  return sheet;
}

// getNextSaturday関数は共通関数.jsのgetCurrentOrNextSaturdayに移行済み

// 日付のメッセージをフォーマット
function formattedDateMessage(date) {
  return (date.getMonth() + 1) + "月" + date.getDate() + "日までの累計時数";
}

// showAlert関数は共通関数.jsに移行済み

// 指定学年の累計結果を計算
function calculateResultsForGrade(data, grade, endDate, categories) {
  const results = {
    "授業時数": 0,
    "儀式": 0,
    "文化": 0,
    "保健": 0,
    "遠足": 0,
    "勤労": 0,
    "欠時数": 0,
    "児童会": 0,
    "クラブ": 0,
    "委員会活動": 0
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateValue = row[DATE_COLUMN];
    if (!dateValue) continue; // 日付がない場合はスキップ

    const date = new Date(dateValue);

    // 日付が無効または集計対象日付を超える場合はスキップ
    if (isNaN(date.getTime()) || date > endDate) continue;

    // 指定学年の行のみを対象にカウント
    if (row[GRADE_COLUMN] == grade) {
      // "○" のカウントを正しく処理
      for (let j = DATA_COLUMNS_START; j <= DATA_COLUMNS_END; j++) {
        if (row[j] == "○") results["授業時数"]++;
      }

      // カテゴリーごとのカウント
      Object.keys(categories).forEach(function (category) {
        for (let j = DATA_COLUMNS_START; j <= DATA_COLUMNS_END; j++) {
          if (row[j] == categories[category]) {
            results[category]++;
          }
        }
      });
    }
  }

  return results;
}
