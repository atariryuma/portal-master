/**
 * @fileoverview 年度更新ファイル作成機能
 * @description 新年度用にスプレッドシートをコピーし、コピー先ファイルの行事データをクリアします。
 *              元ファイルは変更しません。
 */

function copyAndClear() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = getSettingsSheetOrThrow();

    const sourceFile = DriveApp.getFileById(sourceSpreadsheet.getId());
    const folderId = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_DESTINATION_FOLDER_ID).getValue() || '').trim();
    const filename = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_FILE_NAME).getValue() || '').trim();

    if (!filename) {
      ui.alert('エラー: 複製ファイル名（C5）が空です。');
      return;
    }

    let destinationFolder;
    if (folderId) {
      destinationFolder = DriveApp.getFolderById(folderId);
    } else {
      const parentFolders = sourceFile.getParents();
      if (!parentFolders.hasNext()) {
        ui.alert('エラー: コピー先フォルダを取得できません。C7にフォルダIDを設定してください。');
        return;
      }
      destinationFolder = parentFolders.next();
    }

    const copiedFile = sourceFile.makeCopy(filename, destinationFolder);
    const copiedSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    const copiedSheet = copiedSpreadsheet.getSheetByName('年間行事予定表');

    if (!copiedSheet) {
      ui.alert('エラー: コピー先ファイルに「年間行事予定表」シートが見つかりません。');
      return;
    }

    const lastRow = copiedSheet.getLastRow();
    if (lastRow >= 3) {
      copiedSheet.getRange('D3:S' + lastRow).clearContent();   // 校内行事〜その他
      copiedSheet.getRange('U3:AB' + lastRow).clearContent();  // 校時データ〜給食
    }

    ui.alert(
      '年度更新ファイルを作成しました。\n' +
      'コピー先ファイルの行事データをクリアしました。\n' +
      '元ファイルは変更していません。'
    );
  } catch (error) {
    ui.alert('年度更新ファイル作成でエラーが発生しました: ' + error.toString());
  }
}
