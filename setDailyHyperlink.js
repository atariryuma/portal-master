/**
 * @fileoverview 今日の日付へ移動機能
 * @description B1セルに今日の日付へのハイパーリンクを設定します。
 *              日付が見つからない場合はエラーメッセージを表示します。
 */

function setDailyHyperlink() {
  try {
    const sheet = getAnnualScheduleSheet();
    if (!sheet) {
      Logger.log('[ERROR] 年間行事予定表シートが見つかりません');
      return;
    }

    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const dateCol = ANNUAL_SCHEDULE.DATE_COLUMN;
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(dateCol + ANNUAL_SCHEDULE.DATA_START_ROW + ':' + dateCol + lastRow);
    const values = dataRange.getValues();

    let targetRow = null;
    for (let i = 0; i < values.length; i++) {
      const cellDate = values[i][0];
      if (cellDate instanceof Date && Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "yyyy-MM-dd") === formattedToday) {
        targetRow = i + ANNUAL_SCHEDULE.DATA_START_ROW;
        break;
      }
    }

    if (targetRow !== null) {
      const gid = String(sheet.getSheetId());
      const cellRef = dateCol + targetRow;
      const linkFormula = '=HYPERLINK("#gid=' + gid + '&range=' + cellRef + '", "今日へ")';
      sheet.getRange(dateCol + '1').setFormula(linkFormula);
    } else {
      sheet.getRange(dateCol + '1').setValue("今日（" + formattedToday + "）は年間行事予定表に見つかりませんでした。");
    }
  } catch (error) {
    Logger.log('[ERROR] 日付リンク設定中にエラー: ' + error.toString());
  }
}
