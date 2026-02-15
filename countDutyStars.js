/**
 * @fileoverview 休業期間日直カウント機能
 * @description 年間行事予定表R列の「☆」を日直ごとに集計し、日直表E列へ出力します。
 *
 * R列のセル構造（前提）:
 *   1行目: ☆マーク（例: "☆☆"）— ☆の個数がカウント対象
 *   2行目以降: 担当者の名前（名のみ、改行区切り）
 *   例: "☆☆\n太郎\n花子" → 太郎=2, 花子=2
 */
function countStars() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const yearlyScheduleSheet = getAnnualScheduleSheet();
    const dutyRosterSheet = ss.getSheetByName(DUTY_ROSTER_SHEET.NAME);

    if (!yearlyScheduleSheet) {
      showAlert('年間行事予定表シートが見つからないか、データが不完全です。', 'エラー');
      return;
    }

    if (!dutyRosterSheet) {
      showAlert('「日直表」シートが見つかりません。', 'エラー');
      return;
    }

    const dutyCol = ANNUAL_SCHEDULE.DUTY_COLUMN_LETTER;
    const yearlyData = yearlyScheduleSheet.getRange(dutyCol + '1:' + dutyCol + yearlyScheduleSheet.getLastRow()).getValues();
    const dutyRosterRange = dutyRosterSheet.getRange(
      1, DUTY_ROSTER_SHEET.NAME_COLUMN, dutyRosterSheet.getLastRow(), 1
    );
    const dutyRosterData = dutyRosterRange.getValues();

    const outputColumn = dutyRosterSheet.getRange(
      1, DUTY_ROSTER_SHEET.OUTPUT_COLUMN, dutyRosterData.length, 1
    );
    const outputData = outputColumn.getValues();

    const starCounts = {};

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

    for (let i = 1; i < dutyRosterData.length; i++) {
      const fullName = dutyRosterData[i][0];

      if (!fullName || typeof fullName !== 'string' || fullName.trim() === '') {
        continue;
      }

      const firstName = extractFirstName(fullName);
      outputData[i][0] = firstName ? (starCounts[firstName] || 0) : 0;
    }

    outputColumn.setValues(outputData);
  } catch (error) {
    showAlert('☆カウント中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}
