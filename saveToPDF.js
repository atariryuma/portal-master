/**
 * @fileoverview 週報PDF保存機能
 * @description 週報シートをPDF形式で保存し、Google Driveに格納します。
 *              行高さを自動調整し、同名ファイルは自動的に置き換えられます。
 */

function saveToPDF() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    WEEKLY_REPORT.SHEET_NAMES.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        adjustRowHeightsForSheet_(sheet);
        exportSheetToPDF_(sheet);
      } else {
        Logger.log('[WARNING] シート ' + sheetName + ' が見つかりません。');
      }
    });
  } catch (error) {
    showAlert('PDF保存中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}

function adjustRowHeightsForSheet_(sheet) {
  const triggerCell = sheet.getRange(WEEKLY_REPORT.TRIGGER_CELL).getValue();

  // トリガーセルの値に応じて前半・後半の行高さを切り替え
  const hasContent = triggerCell !== '';
  const firstHeight = hasContent ? WEEKLY_REPORT.MAX_HEIGHT : WEEKLY_REPORT.MIN_HEIGHT;
  const secondHeight = hasContent ? WEEKLY_REPORT.MIN_HEIGHT : WEEKLY_REPORT.MAX_HEIGHT;

  sheet.setRowHeights(WEEKLY_REPORT.FIRST_RANGE_START, WEEKLY_REPORT.FIRST_RANGE_COUNT, firstHeight);
  sheet.setRowHeights(WEEKLY_REPORT.SECOND_RANGE_START, WEEKLY_REPORT.SECOND_RANGE_COUNT, secondHeight);
  Logger.log('[INFO] シート【' + sheet.getName() + '】: 行高さを調整しました（前半=' + firstHeight + 'px, 後半=' + secondHeight + 'px）');
}

function exportSheetToPDF_(sheet) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetId = sheet.getSheetId();
    const fileName = createFileName_(sheet);

    const url = preparePdfUrl_(spreadsheet.getId(), sheetId);

    const options = {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    };

    const response = UrlFetchApp.fetch(url, options);
    const blob = response.getBlob().setName(fileName + '.pdf');

    const folderId = getWeeklyReportFolderId();
    const folder = DriveApp.getFolderById(folderId);

    // 新規作成を先に行い、成功後に旧ファイルを削除（作成失敗時のデータ消失を防止）
    const oldFiles = [];
    const existingFiles = folder.getFilesByName(fileName + '.pdf');
    while (existingFiles.hasNext()) {
      oldFiles.push(existingFiles.next());
    }

    folder.createFile(blob);

    oldFiles.forEach(function(file) {
      file.setTrashed(true);
    });
    Logger.log('[INFO] ファイル「' + fileName + '.pdf」をフォルダ「' + folder.getName() + '」に保存しました。');
  } catch (error) {
    Logger.log('[ERROR] PDF出力中にエラー（シート: ' + sheet.getName() + '）: ' + error.toString());
    throw error;
  }
}

function preparePdfUrl_(spreadsheetId, sheetId) {
  const opts = WEEKLY_REPORT.PDF_OPTIONS;
  let url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=pdf&id=' + spreadsheetId;
  url += '&gid=' + sheetId
    + '&size=' + opts.SIZE
    + '&portrait=' + opts.PORTRAIT
    + '&fitw=' + opts.FIT_WIDTH
    + '&top_margin=' + opts.TOP_MARGIN
    + '&right_margin=' + opts.RIGHT_MARGIN
    + '&bottom_margin=' + opts.BOTTOM_MARGIN
    + '&left_margin=' + opts.LEFT_MARGIN
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenum=UNDEFINED'
    + '&scale=' + opts.SCALE
    + '&horizontal_alignment=' + opts.HORIZONTAL_ALIGNMENT
    + '&vertical_alignment=' + opts.VERTICAL_ALIGNMENT
    + '&gridlines=false'
    + '&fzr=false'
    + '&fzc=false';
  return url;
}

function createFileName_(sheet) {
  const range1 = sheet.getRange(WEEKLY_REPORT.NAME_RANGE).getValues()[0].join('');
  const dateRange = sheet.getRange(WEEKLY_REPORT.DATE_RANGE).getValues()[0];
  const formattedDateRange = formatDateToJapanese(dateRange[0]) + '～' + formatDateToJapanese(dateRange[dateRange.length - 1]);
  return range1 + '（' + formattedDateRange + '）';
}
