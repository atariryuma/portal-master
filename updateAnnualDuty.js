function updateAnnualDuty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getSheetByName('マスター');
  const eventSheet = ss.getSheetByName('年間行事予定表');
  
  // A2:AP のデータを取得（必要な列まで取得）
  const masterData = masterSheet.getRange('A2:AP' + masterSheet.getLastRow()).getValues();
  
  // 日付をキーにしたマップを作成（共通関数を使用）
  const dateMap = createDateMap(eventSheet, 'B');
  const totalRows = masterData.length;
  
  // 処理開始を通知
  ui.alert('日直のみの更新を開始します。');
  
  // バッチ更新用の配列を準備
  const dutyUpdates = [];
  
  // マスターデータを日付を基準に「年間行事予定表」へ反映
  masterData.forEach((row, index) => {
    const date = formatDateToJapanese(row[0]); // 共通関数を使用
    const duty = row[40];                      // 日直 (マスターの列AP → 配列index 40)
    if (dateMap[date]) {
      const rowNum = dateMap[date];            // 1-basedの行番号を取得
      dutyUpdates.push({ row: rowNum, value: duty });
      Logger.log(`Processing row ${index + 2}/${totalRows}, Date: ${date}, Duty: ${duty}`);
    }
  });
  
  // バッチで日直を更新
  dutyUpdates.forEach(update => {
    eventSheet.getRange(update.row, 18).setValue(update.value); // R列(列番号18)に設定
  });
  
  // 処理完了を通知
  ui.alert('日直のインポートが完了しました。');
}

// formatDateとcreateDateMap関数は共通関数.jsに移行済み
