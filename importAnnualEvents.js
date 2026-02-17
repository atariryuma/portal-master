/**
 * @fileoverview 年間行事計画インポート機能
 * @description 別スプレッドシートの「メインデータ」シートから、
 *              アクティブなスプレッドシート内の「マスター」シートへ、
 *              対象日（設定シートの基準日=C11セルの日曜日の翌日＝4月1日）から366行分の
 *              値・書式（数値書式、背景色、フォント色、フォントファミリー）を転記します。
 */
function importAnnualEvents() {
  const ui = SpreadsheetApp.getUi();
  try {
    const response = ui.prompt("年間行事計画のインポート",
      "Googleスプレッドシート[Excel小学校年間行事計画（編集用）]のURLを入力してください。",
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    const url = response.getResponseText().trim();
    if (!url || !/^https:\/\/docs\.google\.com\/spreadsheets\/d\//.test(url)) {
      showAlert('GoogleスプレッドシートのURLを入力してください。\n例: https://docs.google.com/spreadsheets/d/...', 'エラー');
      return;
    }

    let sourceSpreadsheet;
    try {
      sourceSpreadsheet = SpreadsheetApp.openByUrl(url);
    } catch (e) {
      showAlert('スプレッドシートを開けませんでした。URLが正しいか、アクセス権限があるか確認してください。', 'エラー');
      return;
    }

    const sourceSheet = sourceSpreadsheet.getSheetByName(IMPORT_CONSTANTS.SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      showAlert('Excel小学校年間行事計画（編集用）に「' + IMPORT_CONSTANTS.SOURCE_SHEET_NAME + '」シートが見つかりません。', 'エラー');
      return;
    }

    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let updateSheet;
    try {
      updateSheet = getSettingsSheetOrThrow();
    } catch (error) {
      showAlert('設定シート（' + SETTINGS_SHEET_NAME + '）が見つかりません。', 'エラー');
      return;
    }
    const sundayDate = normalizeToDate(updateSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.BASE_SUNDAY).getValue());
    if (!sundayDate) {
      showAlert('年度更新設定（C11）に有効な日付が設定されていません。\nシステム管理 → 年度更新設定 で基準日（日曜日）を設定してください。', 'エラー');
      return;
    }

    const year = sundayDate.getFullYear();
    const aprilMonth = MODULE_FISCAL_YEAR_START_MONTH - 1; // 0-based month
    const aprilThisYear = new Date(year, aprilMonth, 1);
    const aprilNextYear = new Date(year + 1, aprilMonth, 1);
    const aprilLastYear = new Date(year - 1, aprilMonth, 1);

    const diffThisYear = Math.abs(sundayDate - aprilThisYear);
    const diffNextYear = Math.abs(sundayDate - aprilNextYear);
    const diffLastYear = Math.abs(sundayDate - aprilLastYear);

    let targetDate = aprilThisYear;
    if (diffNextYear < diffThisYear) {
      targetDate = aprilNextYear;
    }
    if (diffLastYear < Math.abs(sundayDate - targetDate)) {
      targetDate = aprilLastYear;
    }

    const targetDisplayString = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    const sourceValues = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), 1).getValues();
    const sourceStartRow = findDateRow_(sourceValues, targetDisplayString, targetDate.getFullYear());
    if (!sourceStartRow) {
      showAlert('コピー元シートのA列全体に対象の日付 (' + targetDisplayString + ') が見つかりませんでした。', 'エラー');
      return;
    }

    const masterSheet = activeSpreadsheet.getSheetByName(MASTER_SHEET.NAME);
    if (!masterSheet) {
      showAlert('マスターが見つかりません。', 'エラー');
      return;
    }
    const masterValues = masterSheet.getRange(1, 1, masterSheet.getLastRow(), 1).getValues();
    const destStartRow = findDateRow_(masterValues, targetDisplayString, targetDate.getFullYear());
    if (!destStartRow) {
      showAlert('マスターのA列に対象の日付 (' + targetDisplayString + ') が見つかりませんでした。', 'エラー');
      return;
    }

    const numRowsToCopy = IMPORT_CONSTANTS.ROWS_TO_COPY;
    const lastCol = sourceSheet.getLastColumn();
    const sourceAvailableRows = sourceSheet.getLastRow() - sourceStartRow + 1;
    if (sourceAvailableRows < numRowsToCopy) {
      showAlert('コピー元シートのデータ行が不足しています。必要: ' + numRowsToCopy + '行 / 実際: ' + sourceAvailableRows + '行', 'エラー');
      return;
    }

    const sourceRange = sourceSheet.getRange(sourceStartRow, 1, numRowsToCopy, lastCol);
    const dataValues = sourceRange.getValues();
    const dataNumberFormats = sourceRange.getNumberFormats();
    const dataBackgrounds = sourceRange.getBackgrounds();
    const dataFontColors = sourceRange.getFontColors();
    const dataFontFamilies = sourceRange.getFontFamilies();

    const requiredRows = destStartRow + numRowsToCopy - 1;
    if (masterSheet.getMaxRows() < requiredRows) {
      masterSheet.insertRowsAfter(masterSheet.getMaxRows(), requiredRows - masterSheet.getMaxRows());
    }
    if (masterSheet.getMaxColumns() < lastCol) {
      masterSheet.insertColumnsAfter(masterSheet.getMaxColumns(), lastCol - masterSheet.getMaxColumns());
    }
    const destRange = masterSheet.getRange(destStartRow, 1, numRowsToCopy, lastCol);

    destRange.setValues(dataValues);
    destRange.setNumberFormats(dataNumberFormats);
    destRange.setBackgrounds(dataBackgrounds);
    destRange.setFontColors(dataFontColors);
    destRange.setFontFamilies(dataFontFamilies);

    ui.alert("インポートが完了しました。");
  } catch (error) {
    showAlert('年間行事のインポート中にエラーが発生しました: ' + error.toString(), 'エラー');
  }
}

/**
 * セルの値を "yyyy/MM/dd" 形式の文字列に変換
 */
function convertCellValue_(cellValue, year) {
  if (cellValue === null || cellValue === undefined || cellValue === '') {
    return '';
  }
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  if (typeof cellValue === "string") {
    const m = cellValue.match(/^(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      const month = ("0" + m[1]).slice(-2);
      const day = ("0" + m[2]).slice(-2);
      return year + "/" + month + "/" + day;
    }
    return cellValue;
  }
  return cellValue.toString();
}

function findDateRow_(values, targetDisplayString, year) {
  for (let i = 0; i < values.length; i++) {
    const cellString = convertCellValue_(values[i][0], year);
    if (cellString === targetDisplayString) {
      return i + 1;
    }
  }
  return null;
}
