/**
 * @fileoverview 日直のみ更新機能
 * @description マスターのAP列(日直)を、日付一致で年間行事予定表R列に反映します。
 */
function updateAnnualDuty() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const masterSheet = ss.getSheetByName(MASTER_SHEET.NAME);
    const eventSheet = getAnnualScheduleSheetOrThrow();

    if (!masterSheet) {
      showAlert('「マスター」シートが見つかりません。', 'エラー');
      return;
    }

    const masterData = masterSheet.getRange('A' + MASTER_SHEET.DATA_START_ROW + ':' + MASTER_SHEET.DATA_RANGE_END + masterSheet.getLastRow()).getValues();
    const dateMap = createDateMap(eventSheet, ANNUAL_SCHEDULE.DATE_COLUMN);

    ui.alert('日直のみの更新を開始します。');

    const dutyUpdates = [];

    masterData.forEach((row, index) => {
      const date = formatDateToJapanese(row[0]);
      const duty = row[MASTER_SHEET.DUTY_SOURCE_INDEX]; // AP列

      if (dateMap[date]) {
        dutyUpdates.push({ row: dateMap[date], value: duty });
        Logger.log(`[DEBUG] Processing row ${index + 2}/${masterData.length}, Date: ${date}, Duty: ${duty}`);
      }
    });

    dutyUpdates.forEach(update => {
      eventSheet.getRange(update.row, ANNUAL_SCHEDULE.DUTY_COLUMN).setValue(update.value);
    });

    ui.alert('日直のインポートが完了しました。');
  } catch (error) {
    showAlert('日直更新中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}
