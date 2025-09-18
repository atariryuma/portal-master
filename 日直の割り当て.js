function assignDuty() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('日直表');
  var sheet2 = ss.getSheetByName('マスター');

  // 実行を示すダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var processingDialog = ui.alert('実行', '日直表を基にマスターに日直を割り当てます。', ui.ButtonSet.OK);

  // AO2:AO368の範囲をクリア
  sheet2.getRange("AO2:AO368").clearContent();

  // 日直表から日直のペアを取得
  var dutyPairs = {};
  var lastRowSheet1 = sheet1.getLastRow();
  for (var i = 2; i <= lastRowSheet1; i++) {
    var name = sheet1.getRange(i, 3).getValue(); // C列: 氏名
    var dutyNumber = sheet1.getRange(i, 4).getValue(); // D列: 日直番号
    if (name !== "" && dutyNumber !== "") { // 空欄をスキップ
      if (!dutyPairs[dutyNumber]) {
        dutyPairs[dutyNumber] = [];
      }
      dutyPairs[dutyNumber].push(extractFirstName(name)); // 共通関数を使用して名前部分を抽出
    }
  }

  var dutyNumbers = Object.keys(dutyPairs); // 日直番号の配列
  var dutyIndex = 0; // 日直のインデックス
  var dutyPairsLength = dutyNumbers.length;

  // マスターの各行について処理（2行目から始める）
  for (var j = 2; j <= 370; j++) {
    var hasText = false;
    for (var k = 5; k <= 40; k++) { // E列からAN列
      var cellValue = sheet2.getRange(j, k).getValue();
      if (cellValue !== "" && typeof cellValue === "string") {
        hasText = true;
        break;
      }
    }
    // 文字列がある場合、日直を割り当て
    if (hasText) {
      var dutyNumberForThisRow = dutyNumbers[dutyIndex]; // 現在の日直番号
      var namesToAssign = dutyPairs[dutyNumberForThisRow];
      if (namesToAssign && namesToAssign.length > 0) {
        var formattedNames = joinNamesWithNewline(namesToAssign); // 共通関数を使用して改行結合
        Logger.log('Row: ' + j + ', Duty Pair: ' + formattedNames); // ログ出力
        var cell = sheet2.getRange(j, 41); // AO列
        cell.setNumberFormat('@'); // セルの書式をテキスト形式に設定
        cell.setValue(formattedNames);
        cell.setVerticalAlignment('middle').setHorizontalAlignment('center'); // センタリング
      }

      dutyIndex = (dutyIndex + 1) % dutyPairsLength; // 次の日直に移動（ループ）
    }
  }

  // 処理完了メッセージを表示
  ui.alert('完了', '日直の割り当てが完了しました。', ui.ButtonSet.OK);
}
