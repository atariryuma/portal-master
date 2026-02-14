/**
 * @fileoverview 週報PDF保存機能
 * @description 週報シートをPDF形式で保存し、Google Driveに格納します。
 *              行高さを自動調整し、同名ファイルは自動的に置き換えられます。
 */

function saveToPDF() {
  // 対象とするシート名の配列
  const sheetNames = ['週報（時数あり）', '週報（時数あり）次週'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 各シートごとに、行高さの調整とPDF出力を実行
  sheetNames.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      adjustRowHeightsForSheet(sheet);
      exportSheetToPDF(sheet);
    } else {
      Logger.log('シート ' + sheetName + ' が見つかりません。');
    }
  });
}

function adjustRowHeightsForSheet(sheet) {
  // 各シートでセルU41の値を取得
  const triggerCell = sheet.getRange('U41').getValue();  // トリガーセル

  // 行の高さを設定する範囲
  const firstRangeStart = 40; // 40～45行目（6行）
  const firstRangeCount = 6;
  const secondRangeStart = 57; // 57～62行目（6行）
  const secondRangeCount = 6;

  const minHeight = 6;  // 最小高さ
  const maxHeight = 14; // 高い設定値

  if (triggerCell !== '') {
    // U41が空白でない場合
    sheet.setRowHeights(firstRangeStart, firstRangeCount, maxHeight);
    sheet.setRowHeights(secondRangeStart, secondRangeCount, minHeight);
    Logger.log(`シート【${sheet.getName()}】: U41が空白でないため、行${firstRangeStart}-${firstRangeStart + firstRangeCount - 1}を${maxHeight}px, 行${secondRangeStart}-${secondRangeStart + secondRangeCount - 1}を${minHeight}pxに設定しました。`);
  } else {
    // U41が空白の場合
    sheet.setRowHeights(firstRangeStart, firstRangeCount, minHeight);
    sheet.setRowHeights(secondRangeStart, secondRangeCount, maxHeight);
    Logger.log(`シート【${sheet.getName()}】: U41が空白のため、行${firstRangeStart}-${firstRangeStart + firstRangeCount - 1}を${minHeight}px, 行${secondRangeStart}-${secondRangeStart + secondRangeCount - 1}を${maxHeight}pxに設定しました。`);
  }
}

function exportSheetToPDF(sheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = sheet.getSheetId();
  const fileName = createFileName(sheet);

  // PDFエクスポート用のURLを作成
  const url = preparePdfUrl(spreadsheet.getId(), sheetId);

  const options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };

  // URLからPDFを取得し、Blobとして保存
  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob().setName(fileName + '.pdf');

  // 保存先のフォルダIDを取得（共通関数を使用）
  const folderId = getWeeklyReportFolderId();
  const folder = DriveApp.getFolderById(folderId);

  // 同名の既存ファイルを削除
  const files = folder.getFilesByName(fileName + '.pdf');
  while (files.hasNext()) {
    const file = files.next();
    DriveApp.getFileById(file.getId()).setTrashed(true);
  }

  // 新しいPDFファイルを保存
  folder.createFile(blob);
  Logger.log(`ファイル「${fileName}.pdf」をフォルダ「${folder.getName()}」に保存しました。`);
}

function preparePdfUrl(spreadsheetId, sheetId) {
  let url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=pdf&id=' + spreadsheetId;
  // PDFの出力オプションをURLに追加
  url += '&gid=' + sheetId
    + '&size=A4'             // 用紙サイズ (A4)
    + '&portrait=true'       // 縦向き
    + '&fitw=true'           // ページ幅にフィット
    + '&top_margin=0.30'     // 上の余白
    + '&right_margin=0.60'   // 右の余白
    + '&bottom_margin=0.50'  // 下の余白
    + '&left_margin=0.60'    // 左の余白
    + '&sheetnames=false'    // シート名の表示
    + '&printtitle=false'    // スプレッドシート名の表示
    + '&pagenum=UNDEFINED'   // ページ番号の位置
    + '&scale=2'             // 幅に合わせる
    + '&horizontal_alignment=CENTER' // 水平方向中央
    + '&vertical_alignment=CENTER'   // 垂直方向中央
    + '&gridlines=false'     // グリッドライン非表示
    + '&fzr=false'           // 固定行非表示
    + '&fzc=false';          // 固定列非表示
  return url;
}

function createFileName(sheet) {
  // セルB1～D1の値を連結してファイル名の一部とし、セルM1～P1の日付範囲を追加
  const range1 = sheet.getRange('B1:D1').getValues()[0].join('');
  const dateRange = sheet.getRange('M1:P1').getValues()[0];
  const formattedDateRange = formatDateRange(dateRange);
  return range1 + '（' + formattedDateRange + '）';
}

function formatDateRange(dateRange) {
  const start = new Date(dateRange[0]);
  const end = new Date(dateRange[dateRange.length - 1]);
  const formatDate = function(date) {
    return (date.getMonth() + 1) + '月' + date.getDate() + '日';
  };
  return formatDate(start) + '～' + formatDate(end);
}

// getOrCreateFolderId関数はcommon.jsのgetWeeklyReportFolderIdに移行

