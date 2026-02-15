/**
 * @fileoverview 週報フォルダを開く機能
 * @description 週報PDFの保存先フォルダをブラウザで開きます。
 *              フォルダが存在しない場合は自動作成されます。
 */

function openWeeklyReportFolder() {
  try {
    const folderId = getWeeklyReportFolderId();

    // フォルダIDの形式バリデーション
    if (!/^[A-Za-z0-9_-]+$/.test(folderId)) {
      showAlert('フォルダIDの形式が不正です。', 'エラー');
      return;
    }

    const template = HtmlService.createTemplate(
      '<p>週報フォルダを開くには、<a href="https://drive.google.com/drive/folders/<?= encodeURIComponent(folderId) ?>" target="_blank">ここをクリックしてください。</a></p>'
    );
    template.folderId = folderId;

    const htmlOutput = template.evaluate()
      .setWidth(250)
      .setHeight(80);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '週報フォルダを開く');
  } catch (error) {
    showAlert('週報フォルダを開く際にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}
