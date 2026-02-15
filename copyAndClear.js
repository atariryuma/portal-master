/**
 * @fileoverview 年度更新ファイル作成機能
 * @description 新年度用にスプレッドシートをバックアップとしてコピーし、
 *              現在利用中のファイル（URL不変）の行事データをクリアします。
 */

function copyAndClear() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = getSettingsSheetOrThrow();
    const sourceSheet = sourceSpreadsheet.getSheetByName('年間行事予定表');

    if (!sourceSheet) {
      ui.alert('エラー: 現在のファイルに「年間行事予定表」シートが見つかりません。');
      return;
    }

    const sourceFile = DriveApp.getFileById(sourceSpreadsheet.getId());
    const folderId = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_DESTINATION_FOLDER_ID).getValue() || '').trim();
    const filename = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_FILE_NAME).getValue() || '').trim();

    if (!filename) {
      ui.alert('エラー: 複製ファイル名（C5）が空です。');
      return;
    }

    const confirmation = ui.alert(
      '年度更新の確認',
      'バックアップを作成した後、このファイル（現在のURL）の「年間行事予定表」データをクリアします。\n続行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (confirmation !== ui.Button.OK) {
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

    const lastRow = sourceSheet.getLastRow();
    if (lastRow >= 3) {
      sourceSheet.getRange('D3:S' + lastRow).clearContent();   // 校内行事〜その他
      sourceSheet.getRange('U3:AB' + lastRow).clearContent();  // 校時データ〜給食
    }

    ui.alert(
      '年度更新ファイルを作成しました。\n' +
      'バックアップ: ' + copiedFile.getName() + '\n' +
      '現在のファイル（このURL）の行事データをクリアしました。'
    );
  } catch (error) {
    ui.alert('年度更新ファイル作成でエラーが発生しました: ' + error.toString());
  }
}
