function CopyAndClear() {

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
    var folderId = parentFolders.next().getId();    //フォルダのIDを取得
  }


  let folder = DriveApp.getFolderById(folderId);    //フォルダを指定

  var filename = sh1.getRange('C5').getValue();   //作成するファイル名を取得
  

  file.makeCopy(filename, folder);    //コピーの作成

  sh2.getRange('D3:S').clearContent();   //シートのデータを削除
  sh2.getRange('U3:AB').clearContent();   //シートのデータを削除

}


