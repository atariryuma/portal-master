/**
 * @fileoverview 週報フォルダを開く機能
 * @description 週報PDFの保存先フォルダをブラウザで開きます。
 *              フォルダが存在しない場合は自動作成されます。
 */

function openWeeklyReportFolder() {
  // 共通関数を使用してフォルダIDを取得
  var folderId = getWeeklyReportFolderId();
  var folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
  var htmlOutput = HtmlService.createHtmlOutput('<p>週報フォルダを開くには、<a href="' + folderUrl + '" target="_blank">ここをクリックしてください。</a></p>')
      .setWidth(250)
      .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '週報フォルダを開く');
}
