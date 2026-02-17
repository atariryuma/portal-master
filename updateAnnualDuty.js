/**
 * @fileoverview 日直のみ更新機能
 * @description マスターのAO列(日直)を、日付一致で年間行事予定表R列に一括バッチ反映します。
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

    const confirmation = ui.alert('確認', '日直のみの更新を開始します。続行しますか？', ui.ButtonSet.OK_CANCEL);
    if (confirmation !== ui.Button.OK) {
      return;
    }

    // 年間行事予定表のR列を一括読み取り
    const eventLastRow = eventSheet.getLastRow();
    const dutyValues = eventSheet.getRange(1, ANNUAL_SCHEDULE.DUTY_COLUMN, eventLastRow, 1).getValues();

    let updateCount = 0;

    masterData.forEach(function(row) {
      const date = formatDateKey(row[0]);
      const duty = row[MASTER_SHEET.DUTY_SOURCE_INDEX];

      if (dateMap[date]) {
        dutyValues[dateMap[date] - 1][0] = duty;
        updateCount++;
      }
    });

    // 一括書き込み
    if (updateCount > 0) {
      eventSheet.getRange(1, ANNUAL_SCHEDULE.DUTY_COLUMN, eventLastRow, 1).setValues(dutyValues);
    }

    ui.alert('日直のインポートが完了しました。');
  } catch (error) {
    showAlert('日直更新中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}
