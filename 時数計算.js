/**
 * 日付選択ダイアログを表示
 */
function aggregateSchoolEventsByGrade() {
  try {
    var ui = SpreadsheetApp.getUi();
    var htmlOutput = HtmlService.createHtmlOutputFromFile('DateSelector');
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
  var templateSheetName = '時数様式';

  // 学年グループ化: 低(1,2)、中(3,4)、高(5,6)
  var gradeGroups = {
    '低学年': [1, 2],
    '中学年': [3, 4],
    '高学年': [5, 6],
  };

  // 行事カテゴリ(従来通り)
  // 共通関数からカテゴリーを取得
  var categories = EVENT_CATEGORIES;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcSheet = getAnnualScheduleSheet(); // 共通関数を使用
  if (!srcSheet) {
    SpreadsheetApp.getUi().alert('年間行事予定表シートが見つからないか、データが不完全です。');
    return;
  }
  var data = srcSheet.getDataRange().getValues();

  var startDateObj = new Date(startDate);
  var endDateObj = new Date(endDate);

  if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
    SpreadsheetApp.getUi().alert('入力された日付が無効です。');
    return;
  }

  var templateSheet = ss.getSheetByName(templateSheetName);
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert('時数様式シートが見つかりません。');
    return;
  }
  // テンプレートは表示しない
  templateSheet.hideSheet();

  /**
   * gradeGroups でループ: 低学年、中学年、高学年の3パターン
   */
  Object.keys(gradeGroups).forEach(function(groupName) {
    // 例: groupName='低学年'、grades=[1,2]
    var grades = gradeGroups[groupName];

    // 新しいシート名: 例) '低学年'
    var newSheetName = groupName;
    var newSheet = ss.getSheetByName(newSheetName);
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
      // 例えば 0→0行、1→+21行 など
      var blockOffset = idx * 21; // テンプレート次第で調整

      // A2 と A23 のように 21行下にあるとして
      // A2 → row=2, A23 → row=2 + blockOffset(=21)=23
      var gradeCellRow = 2 + blockOffset;
      newSheet.getRange(gradeCellRow, 1).setValue('【' + grade + '年】');

      // 標準授業時数を書き込む例: C17 と C38 (=17+21)
      var standardHourRow = 17 + blockOffset; 
      newSheet.getRange(standardHourRow, 3).setValue(gradeHours[grade]);

      // 集計結果を格納するオブジェクト results
      var results = {};

      // startDateObj から endDateObj まで、1ヶ月ずつ進める
      for (var d = new Date(startDateObj); d <= endDateObj; d.setMonth(d.getMonth() + 1)) {
        var monthKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
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
      }

      // データ集計
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var date = new Date(row[0]);
        if (isNaN(date.getTime())) continue;

        if (date >= startDateObj && date <= endDateObj) {
          var monthKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
          // row[19] が学年: grade(1～6)
          if (row[19] == grade) {
            var hasClass = false;
            // 授業(○)をカウント
            for (var j = 20; j <= 25; j++) {
              if (row[j] == "○") {
                results[monthKey]["授業時数"]++;
                hasClass = true;
              }
            }
            // 行事カテゴリをカウント
            Object.keys(categories).forEach(function(category) {
              for (var j = 20; j <= 25; j++) {
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
      var rowIndexBase = 4 + blockOffset; 
      for (var d2 = new Date(startDateObj); d2 <= endDateObj; d2.setMonth(d2.getMonth() + 1)) {
        var monthKey2 = Utilities.formatDate(d2, 'Asia/Tokyo', 'yyyy-MM');
        if (results[monthKey2]) {
          newSheet.getRange(rowIndexBase, 1).setValue(monthKey2);
          newSheet.getRange(rowIndexBase, 2).setValue(results[monthKey2]["対象日数"]);
          newSheet.getRange(rowIndexBase, 3).setValue(results[monthKey2]["授業時数"]);
          newSheet.getRange(rowIndexBase, 4).setValue(results[monthKey2]["儀式"]);
          newSheet.getRange(rowIndexBase, 5).setValue(results[monthKey2]["文化"]);
          newSheet.getRange(rowIndexBase, 6).setValue(results[monthKey2]["保健"]);
          newSheet.getRange(rowIndexBase, 7).setValue(results[monthKey2]["遠足"]);
          newSheet.getRange(rowIndexBase, 8).setValue(results[monthKey2]["勤労"]);
          // 省略する列がなければ詰める
          newSheet.getRange(rowIndexBase, 10).setValue(results[monthKey2]["欠時数"]);
          newSheet.getRange(rowIndexBase, 11).setValue(results[monthKey2]["児童会"]);
          newSheet.getRange(rowIndexBase, 12).setValue(results[monthKey2]["クラブ"]);
          newSheet.getRange(rowIndexBase, 13).setValue(results[monthKey2]["委員会活動"]);
          newSheet.getRange(rowIndexBase, 14).setValue(results[monthKey2]["補習"]);

          rowIndexBase++;
        }
      }
    });
  });
}
