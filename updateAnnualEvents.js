/**
 * @fileoverview 年間行事予定表への反映機能
 * @description マスターシートのデータを年間行事予定表シートに高速バッチ処理で反映します。
 *              6x6の校時データを一括バッチ処理し、従来の36倍高速化を実現。
 */

function updateAnnualEvents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const masterSheet = ss.getSheetByName('マスター');
    if (!masterSheet) {
      showAlert('「マスター」シートが見つかりません。', 'エラー');
      return;
    }

    const eventSheet = getAnnualScheduleSheetOrThrow(); // 共通関数を使用してエラーハンドリング
    const masterLastRow = masterSheet.getLastRow();
    if (masterLastRow < 2) {
      showAlert('マスターシートに反映対象データがありません。', '通知');
      return;
    }

    const masterData = masterSheet.getRange('A2:AP' + masterLastRow).getValues();
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
  } catch (error) {
    showAlert(error.message || error.toString(), 'エラー');
  }
}

// バッチ更新を実行する関数
function executeBatchUpdate(sheet, updateBatch) {
  if (updateBatch.length === 0) return;

  // 各更新対象をシートに反映
  updateBatch.forEach(update => {
    // 個別の値を設定
    sheet.getRange(update.rowNum, 4).setValue(update.internalEvent); // 校内行事
    sheet.getRange(update.rowNum, 13).setValue(update.externalEvent); // 対外行事
    sheet.getRange(update.rowNum, 27).setValue(update.lunch); // 給食 (AA列)

    // 校時データを6x6の範囲として一括設定
    const attendanceValues = [];
    for (let i = 0; i < 6; i++) {
      const row = [];
      for (let j = 0; j < 6; j++) {
        let value = update.attendance[i * 6 + j];
        // 曜日+学年の組み合わせ（例: 月１、火２）を「○」に変換
        value = /^[月火水木金土日][１-６]$/.test(value) ? '○' : value;
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
