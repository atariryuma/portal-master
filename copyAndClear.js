/**
 * @fileoverview 年度更新ファイル作成機能
 * @description 新年度用にスプレッドシートをバックアップとしてコピーし、
 *              現在利用中のファイル（URL不変）の行事データをクリアします。
 *
 * 設計判断: 「コピー先をクリア」ではなく「現行ファイルをクリア」する理由:
 * - 現行ファイルのURL（ブックマーク・共有リンク）を維持するため
 * - バックアップ（コピー先）は前年度データを完全保持する必要があるため
 * - クリア前にバックアップの整合性検証を行い、失敗時はクリアを中止する
 */

function copyAndClear() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    showAlert('別のユーザーが年度更新を実行中です。しばらく待ってから再度お試しください。', 'エラー');
    return;
  }

  try {
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = getSettingsSheetOrThrow();
    const sourceSheet = sourceSpreadsheet.getSheetByName('年間行事予定表');

    if (!sourceSheet) {
      showAlert('現在のファイルに「年間行事予定表」シートが見つかりません。', 'エラー');
      return;
    }

    const sourceFile = DriveApp.getFileById(sourceSpreadsheet.getId());
    const folderId = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_DESTINATION_FOLDER_ID).getValue() || '').trim();
    const filename = String(settingsSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_FILE_NAME).getValue() || '').trim();

    if (!filename) {
      showAlert('複製ファイル名（C5）が空です。\nシステム管理 → 年度更新設定 から設定してください。', 'エラー');
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
        showAlert('コピー先フォルダを取得できません。C7にフォルダIDを設定してください。', 'エラー');
        return;
      }
      destinationFolder = parentFolders.next();
    }

    const copiedFile = sourceFile.makeCopy(filename, destinationFolder);

    // バックアップ整合性検証: コピー先ファイルが正常に読み取れることを確認
    try {
      const verifiedFile = DriveApp.getFileById(copiedFile.getId());
      if (!verifiedFile || verifiedFile.getName() !== filename) {
        throw new Error('ファイル名の一致を確認できません。');
      }
      const verifiedSs = SpreadsheetApp.openByUrl(copiedFile.getUrl());
      const verifiedSheet = verifiedSs.getSheetByName('年間行事予定表');
      if (!verifiedSheet || verifiedSheet.getLastRow() < 2) {
        throw new Error('コピー先の年間行事予定表シートが不完全です。');
      }
    } catch (verifyError) {
      showAlert(
        'バックアップファイルの検証に失敗しました。データのクリアは実行しません。\n' +
        '手動でバックアップを確認してください。\n詳細: ' + verifyError.toString(),
        'エラー'
      );
      return;
    }

    const lastRow = sourceSheet.getLastRow();
    if (lastRow >= 3) {
      sourceSheet.getRange(ANNUAL_SCHEDULE.CLEAR_EVENT_RANGE + '3:' + ANNUAL_SCHEDULE.CLEAR_EVENT_END + lastRow).clearContent();
      sourceSheet.getRange(ANNUAL_SCHEDULE.CLEAR_DATA_RANGE + '3:' + ANNUAL_SCHEDULE.CLEAR_DATA_END + lastRow).clearContent();
    }

    ui.alert(
      '年度更新ファイルを作成しました。\n' +
      'バックアップ: ' + copiedFile.getName() + '\n' +
      '現在のファイル（このURL）の行事データをクリアしました。'
    );
  } catch (error) {
    showAlert('年度更新ファイル作成でエラーが発生しました: ' + error.toString(), 'エラー');
  } finally {
    lock.releaseLock();
  }
}
