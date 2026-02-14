/**
 * @fileoverview 年度更新ファイル作成機能
 * @description 新年度用にスプレッドシートをコピーし、データをクリアします。
 *              設定値は保持され、行事予定のみリセットされます。
 */

function copyAndClear() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();   //アクティブなスプレッドシートを取得
  var sh1 = ss.getSheetByName('年度更新作業');    //「年度更新作業」シートの指定
  var sh2 = getAnnualScheduleSheet();    // 共通関数を使用

  if (!sh2) {
    SpreadsheetApp.getUi().alert('エラー: 年間行事予定表シートが見つからないか、データが不完全です。');
    return;
  }

  var id = ss.getId();    //スプレッドシートのIDを取得
  var file = DriveApp.getFileById(id);    //コピーするファイルを指定

  var folderId = sh1.getRange('C7').getValue();   //ファイルを格納するフォルダのIDを取得

  if(folderId === ''){    //C7セルが空白の場合
    var parentFolders = DriveApp.getFileById(ss.getId()).getParents();    //このスプレッドシートがあるフォルダを取得
    folderId = parentFolders.next().getId();    //フォルダのIDを取得（var削除）
  }


  let folder = DriveApp.getFolderById(folderId);    //フォルダを指定

  var filename = sh1.getRange('C5').getValue();   //作成するファイル名を取得


  file.makeCopy(filename, folder);    //コピーの作成

  // データクリア範囲を動的に設定（最終行まで）
  var lastRow = sh2.getLastRow();
  sh2.getRange('D3:S' + lastRow).clearContent();   //シートのデータを削除（校内行事〜その他）
  sh2.getRange('U3:AB' + lastRow).clearContent();  //シートのデータを削除（校時データ〜給食）

}

