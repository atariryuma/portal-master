function countStars() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // シートの存在をチェック
  var yearlyScheduleSheet = getAnnualScheduleSheet(); // 共通関数を使用
  var dutyRosterSheet = ss.getSheetByName("日直表");

  if (!yearlyScheduleSheet) {
    SpreadsheetApp.getUi().alert("エラー: 年間行事予定表シートが見つからないか、データが不完全です。");
    return;
  }
  if (!dutyRosterSheet) {
    SpreadsheetApp.getUi().alert("エラー: 「日直表」シートが見つかりません。");
    return;
  }
  
  // 年間行事予定表からデータを取得 (データのある範囲のみ)
  var yearlyData = yearlyScheduleSheet.getRange("R1:R" + yearlyScheduleSheet.getLastRow()).getValues();
  
  // 日直表からデータを取得 (データのある範囲のみ)
  var dutyRosterRange = dutyRosterSheet.getRange("C1:C" + dutyRosterSheet.getLastRow());
  var dutyRosterData = dutyRosterRange.getValues();
  
  var outputColumn = dutyRosterSheet.getRange("E1:E" + dutyRosterData.length);
  var outputData = outputColumn.getValues(); // 既存のデータを取得

  var starCounts = {};
  
  // 年間行事予定表のデータ（下の名前）をキーにして☆をカウント
  for (var i = 0; i < yearlyData.length; i++) {
    var cellContent = yearlyData[i][0];
    if (cellContent && typeof cellContent === 'string') {
      var lines = cellContent.split('\n');
      if (lines.length >= 2) {
        var starCount = (lines[0].match(/☆/g) || []).length;
        if (starCount > 0) {
          for (var j = 1; j < lines.length; j++) {
            var firstName = lines[j].trim(); // ここでは下の名前が取得される
            if (firstName) {
              starCounts[firstName] = (starCounts[firstName] || 0) + starCount;
            }
          }
        }
      }
    }
  }
  
  // 日直表のフルネームから下の名前を抽出し、結果を出力
  for (var i = 1; i < dutyRosterData.length; i++) { // 1行目（タイトル）はスキップ
    var fullName = dutyRosterData[i][0];
    
    if (fullName && typeof fullName === 'string' && fullName.trim() !== "") {
      // 共通関数を使用して名前部分を抽出
      var firstName = extractFirstName(fullName);
      
      if (firstName) {
        var count = starCounts[firstName] || 0; // 下の名前で☆の数を検索
        outputData[i][0] = count;
      } else {
        // フルネームが分割できなかった場合は0とする
        outputData[i][0] = 0;
      }
    }
    // fullNameが空の場合は既存の値を保持
  }
  
  // 結果をE列に一括出力
  outputColumn.setValues(outputData);
}