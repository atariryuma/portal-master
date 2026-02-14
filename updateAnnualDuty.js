/**
 * @fileoverview 日直のみ更新機能
 * @description マスターのAP列(日直)を、日付一致で年間行事予定表R列に反映します。
 */
function updateAnnualDuty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getSheetByName('マスター');
  const eventSheet = getAnnualScheduleSheetOrThrow();

  if (!masterSheet) {
    showAlert('「マスター」シートが見つかりません。', 'エラー');
    return;
  }

  const masterData = masterSheet.getRange('A2:AP' + masterSheet.getLastRow()).getValues();
  const dateMap = createDateMap(eventSheet, 'B');

  ui.alert('日直のみの更新を開始します。');

  const dutyUpdates = [];

  masterData.forEach((row, index) => {
    const date = formatDateToJapanese(row[0]);
    const duty = row[40]; // AP列

    if (dateMap[date]) {
      dutyUpdates.push({ row: dateMap[date], value: duty });
      Logger.log(`Processing row ${index + 2}/${masterData.length}, Date: ${date}, Duty: ${duty}`);
    }
  });

  dutyUpdates.forEach(update => {
    eventSheet.getRange(update.row, 18).setValue(update.value); // R列
  });

  ui.alert('日直のインポートが完了しました。');
}
