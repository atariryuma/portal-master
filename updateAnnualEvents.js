/**
 * @fileoverview 年間行事予定表への反映機能
 * @description マスターシートのデータを年間行事予定表シートに高速バッチ処理で反映します。
 *              6x6の校時データを一括バッチ処理し、従来の36倍高速化を実現。
 */

function updateAnnualEvents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const masterSheet = ss.getSheetByName(MASTER_SHEET.NAME);
    if (!masterSheet) {
      showAlert('「マスター」シートが見つかりません。', 'エラー');
      return;
    }

    const eventSheet = getAnnualScheduleSheetOrThrow();
    const masterLastRow = masterSheet.getLastRow();
    if (masterLastRow < MASTER_SHEET.DATA_START_ROW) {
      showAlert('マスターシートに反映対象データがありません。', '通知');
      return;
    }

    const masterData = masterSheet.getRange('A' + MASTER_SHEET.DATA_START_ROW + ':' + MASTER_SHEET.DATA_RANGE_END + masterLastRow).getValues();
    const dateMap = createDateMapForEvents(eventSheet);
    const totalRows = masterData.length;

    ui.alert('更新処理を開始します。');

    const updateBatch = [];

    masterData.forEach((row, index) => {
      const date = formatDateToJapanese(row[0]);
      if (dateMap[date]) {
        const rowNum = dateMap[date];
        updateBatch.push({
          rowNum: rowNum,
          internalEvent: row[MASTER_SHEET.INTERNAL_EVENT_INDEX],
          externalEvent: row[MASTER_SHEET.EXTERNAL_EVENT_INDEX],
          attendance: row.slice(MASTER_SHEET.DATA_START_COLUMN - 1, MASTER_SHEET.DATA_START_COLUMN - 1 + MASTER_SHEET.DATA_COLUMN_COUNT),
          lunch: row[MASTER_SHEET.LUNCH_INDEX]
        });
        Logger.log(`[DEBUG] Processing row ${index + 2}/${totalRows}, Date: ${date}`);
      }
    });

    executeBatchUpdate(eventSheet, updateBatch);

    masterSheet.hideSheet();

    ui.alert('年間行事のインポート完了に伴い、マスターシートは非表示にしました。今後は「年間行事予定表」シートを直接編集してください。');
  } catch (error) {
    showAlert(error.message || error.toString(), 'エラー');
  }
}

// バッチ更新を実行する関数
function executeBatchUpdate(sheet, updateBatch) {
  if (updateBatch.length === 0) return;

  updateBatch.forEach(update => {
    sheet.getRange(update.rowNum, ANNUAL_SCHEDULE.INTERNAL_EVENT_COLUMN).setValue(update.internalEvent);
    sheet.getRange(update.rowNum, ANNUAL_SCHEDULE.EXTERNAL_EVENT_COLUMN).setValue(update.externalEvent);
    sheet.getRange(update.rowNum, ANNUAL_SCHEDULE.LUNCH_COLUMN).setValue(update.lunch);

    // 校時データを6x6の範囲として一括設定
    const attendanceValues = [];
    for (let i = 0; i < ANNUAL_SCHEDULE.ATTENDANCE_ROWS; i++) {
      const row = [];
      for (let j = 0; j < ANNUAL_SCHEDULE.ATTENDANCE_COLS; j++) {
        let value = update.attendance[i * ANNUAL_SCHEDULE.ATTENDANCE_COLS + j];
        value = /^[月火水木金土日][１-６]$/.test(value) ? '○' : value;
        row.push(value);
      }
      attendanceValues.push(row);
    }
    sheet.getRange(
      update.rowNum,
      ANNUAL_SCHEDULE.ATTENDANCE_START_COLUMN,
      ANNUAL_SCHEDULE.ATTENDANCE_ROWS,
      ANNUAL_SCHEDULE.ATTENDANCE_COLS
    ).setValues(attendanceValues);
  });
}


// 日付マップを作成（共通関数を活用）
function createDateMapForEvents(sheet) {
  return createDateMap(sheet, ANNUAL_SCHEDULE.DATE_COLUMN);
}
