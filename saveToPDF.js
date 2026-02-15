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
        adjustRowHeightsForSheet(sheet);
        exportSheetToPDF(sheet);
      } else {
        Logger.log('[WARNING] シート ' + sheetName + ' が見つかりません。');
      }
    });
  } catch (error) {
    showAlert('PDF保存中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}

function adjustRowHeightsForSheet(sheet) {
  const triggerCell = sheet.getRange(WEEKLY_REPORT.TRIGGER_CELL).getValue();

  if (triggerCell !== '') {
    sheet.setRowHeights(WEEKLY_REPORT.FIRST_RANGE_START, WEEKLY_REPORT.FIRST_RANGE_COUNT, WEEKLY_REPORT.MAX_HEIGHT);
    sheet.setRowHeights(WEEKLY_REPORT.SECOND_RANGE_START, WEEKLY_REPORT.SECOND_RANGE_COUNT, WEEKLY_REPORT.MIN_HEIGHT);
    Logger.log(`[INFO] シート【${sheet.getName()}】: ${WEEKLY_REPORT.TRIGGER_CELL}が空白でないため、行${WEEKLY_REPORT.FIRST_RANGE_START}-${WEEKLY_REPORT.FIRST_RANGE_START + WEEKLY_REPORT.FIRST_RANGE_COUNT - 1}を${WEEKLY_REPORT.MAX_HEIGHT}px, 行${WEEKLY_REPORT.SECOND_RANGE_START}-${WEEKLY_REPORT.SECOND_RANGE_START + WEEKLY_REPORT.SECOND_RANGE_COUNT - 1}を${WEEKLY_REPORT.MIN_HEIGHT}pxに設定しました。`);
  } else {
    sheet.setRowHeights(WEEKLY_REPORT.FIRST_RANGE_START, WEEKLY_REPORT.FIRST_RANGE_COUNT, WEEKLY_REPORT.MIN_HEIGHT);
    sheet.setRowHeights(WEEKLY_REPORT.SECOND_RANGE_START, WEEKLY_REPORT.SECOND_RANGE_COUNT, WEEKLY_REPORT.MAX_HEIGHT);
    Logger.log(`[INFO] シート【${sheet.getName()}】: ${WEEKLY_REPORT.TRIGGER_CELL}が空白のため、行${WEEKLY_REPORT.FIRST_RANGE_START}-${WEEKLY_REPORT.FIRST_RANGE_START + WEEKLY_REPORT.FIRST_RANGE_COUNT - 1}を${WEEKLY_REPORT.MIN_HEIGHT}px, 行${WEEKLY_REPORT.SECOND_RANGE_START}-${WEEKLY_REPORT.SECOND_RANGE_START + WEEKLY_REPORT.SECOND_RANGE_COUNT - 1}を${WEEKLY_REPORT.MAX_HEIGHT}pxに設定しました。`);
  }
}

function exportSheetToPDF(sheet) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetId = sheet.getSheetId();
    const fileName = createFileName(sheet);

    const url = preparePdfUrl(spreadsheet.getId(), sheetId);

    const options = {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    };

    const response = UrlFetchApp.fetch(url, options);
    const blob = response.getBlob().setName(fileName + '.pdf');

    const folderId = getWeeklyReportFolderId();
    const folder = DriveApp.getFolderById(folderId);

    const files = folder.getFilesByName(fileName + '.pdf');
    while (files.hasNext()) {
      const file = files.next();
      DriveApp.getFileById(file.getId()).setTrashed(true);
    }

    folder.createFile(blob);
    Logger.log(`[INFO] ファイル「${fileName}.pdf」をフォルダ「${folder.getName()}」に保存しました。`);
  } catch (error) {
    Logger.log(`[ERROR] PDF出力中にエラー（シート: ${sheet.getName()}）: ${error.toString()}`);
    throw error;
  }
}

function preparePdfUrl(spreadsheetId, sheetId) {
  let url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=pdf&id=' + spreadsheetId;
  url += '&gid=' + sheetId
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    + '&top_margin=0.30'
    + '&right_margin=0.60'
    + '&bottom_margin=0.50'
    + '&left_margin=0.60'
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenum=UNDEFINED'
    + '&scale=2'
    + '&horizontal_alignment=CENTER'
    + '&vertical_alignment=CENTER'
    + '&gridlines=false'
    + '&fzr=false'
    + '&fzc=false';
  return url;
}

function createFileName(sheet) {
  const range1 = sheet.getRange(WEEKLY_REPORT.NAME_RANGE).getValues()[0].join('');
  const dateRange = sheet.getRange(WEEKLY_REPORT.DATE_RANGE).getValues()[0];
  const formattedDateRange = formatDateRangeForPdf_(dateRange);
  return range1 + '（' + formattedDateRange + '）';
}

function formatDateRangeForPdf_(dateRange) {
  const start = new Date(dateRange[0]);
  const end = new Date(dateRange[dateRange.length - 1]);
  return formatDateToJapanese(start) + '～' + formatDateToJapanese(end);
}
