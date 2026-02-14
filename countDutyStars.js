/**
 * @fileoverview 休業期間日直カウント機能
 * @description 年間行事予定表R列の「☆」を日直ごとに集計し、日直表E列へ出力します。
 */
function countStars() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yearlyScheduleSheet = getAnnualScheduleSheet();
  const dutyRosterSheet = ss.getSheetByName('日直表');

  if (!yearlyScheduleSheet) {
    showAlert('年間行事予定表シートが見つからないか、データが不完全です。', 'エラー');
    return;
  }

  if (!dutyRosterSheet) {
    showAlert('「日直表」シートが見つかりません。', 'エラー');
    return;
  }

  const yearlyData = yearlyScheduleSheet.getRange('R1:R' + yearlyScheduleSheet.getLastRow()).getValues();
  const dutyRosterRange = dutyRosterSheet.getRange('C1:C' + dutyRosterSheet.getLastRow());
  const dutyRosterData = dutyRosterRange.getValues();

  const outputColumn = dutyRosterSheet.getRange('E1:E' + dutyRosterData.length);
  const outputData = outputColumn.getValues();

  const starCounts = {};

  // R列セル: 1行目=☆行、2行目以降=担当者名 の前提で集計
  for (let i = 0; i < yearlyData.length; i++) {
    const cellContent = yearlyData[i][0];

    if (!cellContent || typeof cellContent !== 'string') {
      continue;
    }

    const lines = cellContent.split('\n');
    if (lines.length < 2) {
      continue;
    }

    const starCount = (lines[0].match(/☆/g) || []).length;
    if (starCount === 0) {
      continue;
    }

    for (let j = 1; j < lines.length; j++) {
      const firstName = lines[j].trim();
      if (!firstName) {
        continue;
      }

      starCounts[firstName] = (starCounts[firstName] || 0) + starCount;
    }
  }

  // 日直表の氏名(C列)から下の名前を抽出してE列に出力
  for (let i = 1; i < dutyRosterData.length; i++) {
    const fullName = dutyRosterData[i][0];

    if (!fullName || typeof fullName !== 'string' || fullName.trim() === '') {
      continue;
    }

    const firstName = extractFirstName(fullName);
    outputData[i][0] = firstName ? (starCounts[firstName] || 0) : 0;
  }

  outputColumn.setValues(outputData);
}
