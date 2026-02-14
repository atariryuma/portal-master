/**
 * @fileoverview 今日の日付へ移動機能
 * @description B1セルに今日の日付へのハイパーリンクを設定します。
 *              日付が見つからない場合はエラーメッセージを表示します。
 */

function setDailyHyperlink() {
  const sheet = getAnnualScheduleSheet(); // 共通関数を使用
  if (!sheet) {
    Logger.log('[ERROR] 年間行事予定表シートが見つかりません');
    return;
  }

  const today = new Date();
  const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const dataRange = sheet.getRange("B2:B");
  const values = dataRange.getValues();

  let targetRow = null;
  for (let i = 0; i < values.length; i++) {
    const cellDate = values[i][0];
    if (cellDate instanceof Date && Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "yyyy-MM-dd") === formattedToday) {
      targetRow = i + 2;  // B列のインデックスは1ベースなので2を足す
      break;
    }
  }

  if (targetRow !== null) {
    const hyperlink = `#gid=${sheet.getSheetId()}&range=B${targetRow}`;
    const linkFormula = `=HYPERLINK("${hyperlink}", "今日へ")`;
    sheet.getRange('B1').setFormula(linkFormula);
  } else {
    // 対応する日付が見つからない場合、エラーメッセージを表示
    sheet.getRange('B1').setValue("今日の日付は見つかりませんでした。");
  }
}