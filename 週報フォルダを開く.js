function openWeeklyReportFolder() {
  // 共通関数を使用してフォルダIDを取得
  var folderId = getWeeklyReportFolderId();
  var folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
  var htmlOutput = HtmlService.createHtmlOutput('<p>週報フォルダを開くには、<a href="' + folderUrl + '" target="_blank">ここをクリックしてください。</a></p>')
      .setWidth(250)
      .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '週報フォルダを開く');
}
