/**
 * @fileoverview 日直割り当て機能
 * @description 日直表の番号順に、マスターシートへ日直を割り当てます。
 */
function assignDuty() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const dutySheet = ss.getSheetByName(DUTY_ROSTER_SHEET.NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET.NAME);

    if (!dutySheet || !masterSheet) {
      showAlert('日直表またはマスターシートが見つかりません。', 'エラー');
      return;
    }

    const lastRow = dutySheet.getLastRow();
    if (lastRow < DUTY_ROSTER_SHEET.DATA_START_ROW) {
      showAlert('日直表に割り当て可能なデータがありません。', '通知');
      return;
    }

    // バッチ読み取り: C列・D列を一括取得
    const dutyRosterData = dutySheet.getRange(
      DUTY_ROSTER_SHEET.DATA_START_ROW,
      DUTY_ROSTER_SHEET.NAME_COLUMN,
      lastRow - DUTY_ROSTER_SHEET.DATA_START_ROW + 1,
      DUTY_ROSTER_SHEET.NUMBER_COLUMN - DUTY_ROSTER_SHEET.NAME_COLUMN + 1
    ).getValues();

    const dutyPairs = {};
    for (let i = 0; i < dutyRosterData.length; i++) {
      const fullName = dutyRosterData[i][0]; // C列: 氏名
      const dutyNumber = dutyRosterData[i][1]; // D列: 日直番号

      if (!isNonEmptyCell(fullName) || !isNonEmptyCell(dutyNumber)) {
        continue;
      }

      if (!dutyPairs[dutyNumber]) {
        dutyPairs[dutyNumber] = [];
      }

      dutyPairs[dutyNumber].push(extractFirstName(fullName));
    }

    const dutyNumbers = Object.keys(dutyPairs);
    if (dutyNumbers.length === 0) {
      showAlert('日直表に割り当て可能なデータがありません。', '通知');
      return;
    }

    const confirmation = ui.alert('確認', '日直表を基にマスターへ日直を割り当てます。続行しますか？', ui.ButtonSet.OK_CANCEL);
    if (confirmation !== ui.Button.OK) {
      return;
    }

    // AO列の日直欄を一度クリア
    const endRow = Math.min(MASTER_SHEET.MAX_DATA_ROW, masterSheet.getMaxRows());
    masterSheet.getRange('AO2:AO' + endRow).clearContent();

    // バッチ読み取り: E:AN列を一括取得
    const masterData = masterSheet.getRange(
      MASTER_SHEET.DATA_START_ROW,
      MASTER_SHEET.DATA_START_COLUMN,
      endRow - MASTER_SHEET.DATA_START_ROW + 1,
      MASTER_SHEET.DATA_COLUMN_COUNT
    ).getValues();

    // 出力用配列を構築
    const outputData = [];
    let dutyIndex = 0;
    for (let i = 0; i < masterData.length; i++) {
      const rowValues = masterData[i];
      const hasText = rowValues.some(function(value) { return isNonEmptyCell(value); });

      if (!hasText) {
        outputData.push(['']);
        continue;
      }

      const dutyNumber = dutyNumbers[dutyIndex];
      const namesToAssign = dutyPairs[dutyNumber];

      if (namesToAssign && namesToAssign.length > 0) {
        outputData.push([joinNamesWithNewline(namesToAssign)]);
      } else {
        outputData.push(['']);
      }

      dutyIndex = (dutyIndex + 1) % dutyNumbers.length;
    }

    // バッチ書き込み: AO列に一括設定
    const outputRange = masterSheet.getRange(
      MASTER_SHEET.DATA_START_ROW,
      MASTER_SHEET.DUTY_COLUMN,
      outputData.length,
      1
    );
    outputRange.setNumberFormat('@');
    outputRange.setValues(outputData);
    outputRange.setVerticalAlignment('middle').setHorizontalAlignment('center');

    ui.alert('完了', '日直の割り当てが完了しました。', ui.ButtonSet.OK);
  } catch (error) {
    showAlert('日直割り当て中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}
