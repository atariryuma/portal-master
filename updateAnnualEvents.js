/**
 * @fileoverview 年間行事予定表への反映機能
 * @description マスターシートのデータを年間行事予定表シートに一括バッチ処理で反映します。
 *              全更新をメモリ上で構築し、setValues() 1回で書き込みます。
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

    const confirmation = ui.alert('確認', '年間行事予定表への更新処理を開始します。続行しますか？', ui.ButtonSet.OK_CANCEL);
    if (confirmation !== ui.Button.OK) {
      return;
    }

    // 年間行事予定表の対象列範囲を一括読み取り
    const eventLastRow = eventSheet.getLastRow();
    const eventInternalCol = ANNUAL_SCHEDULE.INTERNAL_EVENT_COLUMN;
    const eventExternalCol = ANNUAL_SCHEDULE.EXTERNAL_EVENT_COLUMN;
    const eventAttStartCol = ANNUAL_SCHEDULE.ATTENDANCE_START_COLUMN;
    const eventLunchCol = ANNUAL_SCHEDULE.LUNCH_COLUMN;

    // 校内行事列(D)を一括取得・更新
    const internalValues = eventSheet.getRange(1, eventInternalCol, eventLastRow, 1).getValues();
    // 対外行事列(M)を一括取得・更新
    const externalValues = eventSheet.getRange(1, eventExternalCol, eventLastRow, 1).getValues();
    // 給食列(AA)を一括取得・更新
    const lunchValues = eventSheet.getRange(1, eventLunchCol, eventLastRow, 1).getValues();
    // 校時データ(U:Z, 6列)を一括取得・更新
    const attendanceValues = eventSheet.getRange(1, eventAttStartCol, eventLastRow, ANNUAL_SCHEDULE.ATTENDANCE_COLS).getValues();

    let updateCount = 0;

    masterData.forEach(function(row) {
      const date = formatDateToJapanese(row[0]);
      if (!dateMap[date]) {
        return;
      }

      const eventRowIndex = dateMap[date] - 1; // 0-based index
      internalValues[eventRowIndex][0] = row[MASTER_SHEET.INTERNAL_EVENT_INDEX];
      externalValues[eventRowIndex][0] = row[MASTER_SHEET.EXTERNAL_EVENT_INDEX];
      lunchValues[eventRowIndex][0] = row[MASTER_SHEET.LUNCH_INDEX];

      // 校時データ: マスターの36列(E:AN)から6行分を抽出し、6x1の行として書き込み
      // マスター1行 = 年間行事予定表の6行(学年行) x 6列(校時列)
      const masterAttendance = row.slice(MASTER_SHEET.DATA_START_COLUMN - 1, MASTER_SHEET.DATA_START_COLUMN - 1 + MASTER_SHEET.DATA_COLUMN_COUNT);
      for (let j = 0; j < ANNUAL_SCHEDULE.ATTENDANCE_COLS; j++) {
        let value = masterAttendance[j];
        value = /^[月火水木金土日][１-６]$/.test(value) ? '○' : value;
        attendanceValues[eventRowIndex][j] = value;
      }

      updateCount++;
    });

    // 一括書き込み
    if (updateCount > 0) {
      eventSheet.getRange(1, eventInternalCol, eventLastRow, 1).setValues(internalValues);
      eventSheet.getRange(1, eventExternalCol, eventLastRow, 1).setValues(externalValues);
      eventSheet.getRange(1, eventLunchCol, eventLastRow, 1).setValues(lunchValues);
      eventSheet.getRange(1, eventAttStartCol, eventLastRow, ANNUAL_SCHEDULE.ATTENDANCE_COLS).setValues(attendanceValues);
    }

    masterSheet.hideSheet();

    ui.alert('年間行事のインポート完了に伴い、マスターシートは非表示にしました。今後は「年間行事予定表」シートを直接編集してください。');
  } catch (error) {
    showAlert(error.message || error.toString(), 'エラー');
  }
}

function createDateMapForEvents(sheet) {
  return createDateMap(sheet, ANNUAL_SCHEDULE.DATE_COLUMN);
}
