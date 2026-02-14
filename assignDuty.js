/**
 * @fileoverview 日直割り当て機能
 * @description 日直表の番号順に、マスターシートへ日直を割り当てます。
 */
function assignDuty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const dutySheet = ss.getSheetByName('日直表');
  const masterSheet = ss.getSheetByName('マスター');

  if (!dutySheet || !masterSheet) {
    showAlert('日直表またはマスターシートが見つかりません。', 'エラー');
    return;
  }

  const dutyPairs = {};
  const lastRow = dutySheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const fullName = dutySheet.getRange(row, 3).getValue(); // C列: 氏名
    const dutyNumber = dutySheet.getRange(row, 4).getValue(); // D列: 日直番号

    if (fullName === '' || dutyNumber === '') {
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

  ui.alert('実行', '日直表を基にマスターへ日直を割り当てます。', ui.ButtonSet.OK);

  // AO列の日直欄を一度クリア（処理対象の最終行まで）
  const endRow = Math.min(370, masterSheet.getMaxRows());
  masterSheet.getRange('AO2:AO' + endRow).clearContent();

  let dutyIndex = 0;
  for (let row = 2; row <= endRow; row++) {
    const rowValues = masterSheet.getRange(row, 5, 1, 36).getValues()[0]; // E:AN
    const hasText = rowValues.some(value => typeof value === 'string' && value !== '');

    if (!hasText) {
      continue;
    }

    const dutyNumber = dutyNumbers[dutyIndex];
    const namesToAssign = dutyPairs[dutyNumber];

    if (namesToAssign && namesToAssign.length > 0) {
      const formattedNames = joinNamesWithNewline(namesToAssign);
      const cell = masterSheet.getRange(row, 41); // AO列
      cell.setNumberFormat('@');
      cell.setValue(formattedNames);
      cell.setVerticalAlignment('middle').setHorizontalAlignment('center');
    }

    dutyIndex = (dutyIndex + 1) % dutyNumbers.length;
  }

  ui.alert('完了', '日直の割り当てが完了しました。', ui.ButtonSet.OK);
}

