function updateAnnualEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getSheetByName('マスター');
  const eventSheet = getAnnualScheduleSheetOrThrow(); // 共通関数を使用してエラーハンドリング
  const masterData = masterSheet.getRange('A2:AP' + masterSheet.getLastRow()).getValues();
  const dateMap = createDateMapForEvents(eventSheet);
  const totalRows = masterData.length;

  // 処理開始を通知
  ui.alert('更新処理を開始します。');

  // バッチ処理用の配列を準備
  const updateBatch = [];

  masterData.forEach((row, index) => {
    const date = formatDateToJapanese(row[0]); // 共通関数を使用
    if (dateMap[date]) {
      const rowNum = dateMap[date];
      updateBatch.push({
        rowNum: rowNum,
        internalEvent: row[2],
        externalEvent: row[3],
        attendance: row.slice(4, 40),
        duty: row[40],
        lunch: row[41]
      });
      Logger.log(`Processing row ${index + 2}/${totalRows}, Date: ${date}`);
    }
  });

  // バッチで更新を実行
  executeBatchUpdate(eventSheet, updateBatch);

  // マスターシートを非表示に設定
  masterSheet.hideSheet();

  // 処理完了を通知
  ui.alert('年間行事のインポート完了に伴い、マスターシートは非表示にしました。今後は「年間行事予定」シートを直接編集してください。');
}

// バッチ更新を実行する関数
function executeBatchUpdate(sheet, updateBatch) {
  if (updateBatch.length === 0) return;

  // 各更新対象をシートに反映
  updateBatch.forEach(update => {
    // 個別の値を設定
    sheet.getRange(update.rowNum, 4).setValue(update.internalEvent); // 校内行事
    sheet.getRange(update.rowNum, 13).setValue(update.externalEvent); // 対外行事
    sheet.getRange(update.rowNum, 18).setValue(update.duty); // 日直 (R列)
    sheet.getRange(update.rowNum, 27).setValue(update.lunch); // 給食 (AA列)

    // 校時データを6x6の範囲として一括設定
    const attendanceValues = [];
    for (let i = 0; i < 6; i++) {
      const row = [];
      for (let j = 0; j < 6; j++) {
        let value = update.attendance[i * 6 + j];
        value = /^[月火水木金土日]１$|^[月火水木金土日]２$|^[月火水木金土日]３$|^[月火水木金土日]４$|^[月火水木金土日]５$|^[月火水木金土日]６$/.test(value) ? '○' : value;
        row.push(value);
      }
      attendanceValues.push(row);
    }
    // 6x6の範囲を一括で設定
    sheet.getRange(update.rowNum, 21, 6, 6).setValues(attendanceValues);
  });
}


// 日付マップを作成（共通関数を活用）
function createDateMapForEvents(sheet) {
  return createDateMap(sheet, 'B');
}
