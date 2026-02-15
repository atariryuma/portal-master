/**
 * 別スプレッドシートの「メインデータ」シートから、  
 * アクティブなスプレッドシート内の「マスター」シートへ、  
 * 対象日（設定シートの基準日=C11セルの日曜日の翌日＝4月1日）から366行分の  
 * 値・書式（数値書式、背景色、フォント色、フォントファミリー）を転記するスクリプト。
 *
 * ※コピー元シートのA列は「4月1日」などと表示されている場合があるため、  
 *    補完する年（対象日から取得）を用いて "yyyy/MM/dd" 形式に変換し、  
 *    対象日（例："2025/04/01"）と比較します。
 *
 * 【デバッグ用ログ】  
 * ・targetDisplayString … 対象日を "yyyy/MM/dd" 形式にした文字列  
 * ・sourceSheet A列上位10行 (変換後) … コピー元シートのA列（1～10行）の値（変換後）  
 * ・masterSheet A列上位10行 (変換後) … 貼り付け先シートのA列（1～10行）の値（変換後）  
 */
function importAnnualEvents() {
  var ui = SpreadsheetApp.getUi();

  // 1. ユーザーに、元スプレッドシートのURLを入力してもらう
  var response = ui.prompt("年間行事計画のインポート", 
    "Googleスプレッドシート[Excel小学校年間行事計画（編集用）]のURLを入力してください。", 
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // キャンセル時は何もしない
  }
  var url = response.getResponseText().trim();
  
  // 2. 入力URLからコピー元のスプレッドシートを取得
  var sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openByUrl(url);
  } catch(e) {
    ui.alert("無効なURLです。スプレッドシートを開けませんでした。");
    return;
  }
  
  // コピー元スプレッドシートの「メインデータ」シートを取得
  var sourceSheet = sourceSpreadsheet.getSheetByName("メインデータ");
  if (!sourceSheet) {
    ui.alert("Excel小学校年間行事計画（編集用）に「メインデータ」シートが見つかりません。");
    return;
  }
  
  // 3. アクティブなスプレッドシート（ポータルマスター）の設定シート（C11）から、
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var updateSheet;
  try {
    updateSheet = getSettingsSheetOrThrow();
  } catch (error) {
    ui.alert("設定シート（" + SETTINGS_SHEET_NAME + "）が見つかりません。");
    return;
  }
  var sundayDate = updateSheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.BASE_SUNDAY).getValue();
  if (!(sundayDate instanceof Date)) {
    sundayDate = new Date(sundayDate);
    if (isNaN(sundayDate.getTime())) {
      ui.alert("年度更新設定（C11）に有効な日付が設定されていません。");
      return;
    }
  }

  // 4月1日の候補を用意（同年、翌年、前年度）
  var year = sundayDate.getFullYear();
  var aprilThisYear = new Date(year, 3, 1);    // 月は0始まりなので、3は4月
  var aprilNextYear = new Date(year + 1, 3, 1);
  var aprilLastYear = new Date(year - 1, 3, 1);

  // 候補との日数差を絶対値で計算
  var diffThisYear = Math.abs(sundayDate - aprilThisYear);
  var diffNextYear = Math.abs(sundayDate - aprilNextYear);
  var diffLastYear = Math.abs(sundayDate - aprilLastYear);

  // 最も近い4月1日を選択
  var targetDate = aprilThisYear;
  if (diffNextYear < diffThisYear) {
    targetDate = aprilNextYear;
  }
  if (diffLastYear < Math.abs(sundayDate - targetDate)) {
    targetDate = aprilLastYear;
  }

  // 対象日を "yyyy/MM/dd" 形式の文字列に変換（例："2025/04/01"）
  var targetDisplayString = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  Logger.log("targetDisplayString: " + targetDisplayString);
  
  // 4. コピー元シートのA列上位10行の【表示値】（そのまま）を取得してログ出力（参考）
  var sourceDisplayLogRowCount = Math.min(10, sourceSheet.getLastRow());
  var sourceDispValues = sourceSheet.getRange(1, 1, sourceDisplayLogRowCount, 1).getDisplayValues();
  Logger.log("sourceSheet A列上位10行 (getDisplayValues): " + JSON.stringify(sourceDispValues));
  
  // 5. コピー元シートA列全体を検索し、対象日がある開始行を決定
  var sourceValues = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), 1).getValues();
  var sourceLogConverted = [];
  for (var i = 0; i < Math.min(10, sourceValues.length); i++) {
    sourceLogConverted.push(convertCellValue(sourceValues[i][0], targetDate.getFullYear()));
  }
  var sourceStartRow = findDateRow(sourceValues, targetDisplayString, targetDate.getFullYear());
  Logger.log("sourceSheet A列上位10行 (converted): " + JSON.stringify(sourceLogConverted));
  if (!sourceStartRow) {
    ui.alert("コピー元シートのA列全体に対象の日付 (" + targetDisplayString + ") が見つかりませんでした。");
    return;
  }
  Logger.log("コピー元対象開始行: " + sourceStartRow);
  
  // 6. 貼り付け先シート「マスター」のA列全体を検索し、対象日がある行番号を決定する
  var masterSheet = activeSpreadsheet.getSheetByName("マスター");
  if (!masterSheet) {
    ui.alert("マスターが見つかりません。");
    return;
  }
  var masterValues = masterSheet.getRange(1, 1, masterSheet.getLastRow(), 1).getValues();
  var masterLogConverted = [];
  for (var j = 0; j < Math.min(10, masterValues.length); j++) {
    masterLogConverted.push(convertCellValue(masterValues[j][0], targetDate.getFullYear()));
  }
  var destStartRow = findDateRow(masterValues, targetDisplayString, targetDate.getFullYear());
  Logger.log("masterSheet A列上位10行 (converted): " + JSON.stringify(masterLogConverted));
  if (!destStartRow) {
    ui.alert("マスターのA列に対象の日付 (" + targetDisplayString + ") が見つかりませんでした。");
    return;
  }
  Logger.log("貼り付け先対象開始行: " + destStartRow);
  
  // 7. コピーする行数（366行）と、コピー元シートの最終列を決定
  var numRowsToCopy = 366;
  var lastCol = sourceSheet.getLastColumn();
  var sourceAvailableRows = sourceSheet.getLastRow() - sourceStartRow + 1;
  if (sourceAvailableRows < numRowsToCopy) {
    ui.alert("コピー元シートのデータ行が不足しています。必要: " + numRowsToCopy + "行 / 実際: " + sourceAvailableRows + "行");
    return;
  }
  
  // コピー元シートの対象範囲（sourceStartRow ～ sourceStartRow+365行、全列）の値と書式情報を取得
  var sourceRange = sourceSheet.getRange(sourceStartRow, 1, numRowsToCopy, lastCol);
  var dataValues = sourceRange.getValues();
  var dataNumberFormats = sourceRange.getNumberFormats();
  var dataBackgrounds = sourceRange.getBackgrounds();
  var dataFontColors = sourceRange.getFontColors();
  var dataFontFamilies = sourceRange.getFontFamilies();
  
  // 8. 貼り付け先シートに、貼り付け先の範囲（destStartRowから366行分）を確保（足りなければ行を追加）
  var requiredRows = destStartRow + numRowsToCopy - 1;
  if (masterSheet.getMaxRows() < requiredRows) {
    masterSheet.insertRowsAfter(masterSheet.getMaxRows(), requiredRows - masterSheet.getMaxRows());
  }
  var destRange = masterSheet.getRange(destStartRow, 1, numRowsToCopy, lastCol);
  
  // 9. コピー元の【値】および【書式情報】を貼り付け先に設定する
  destRange.setValues(dataValues);
  destRange.setNumberFormats(dataNumberFormats);
  destRange.setBackgrounds(dataBackgrounds);
  destRange.setFontColors(dataFontColors);
  destRange.setFontFamilies(dataFontFamilies);
  
  ui.alert("インポートが完了しました。");
}

/**
 * セルの値（生データ）を "yyyy/MM/dd" 形式の文字列に変換するヘルパー関数  
 * ※ cellValue が Date オブジェクトの場合はそのままフォーマット、  
 *    文字列で "X月Y日" 形式ならば、year を補完して "yyyy/MM/dd" 形式に変換します。
 */
function convertCellValue(cellValue, year) {
  var result;
  if (cellValue === null || cellValue === undefined || cellValue === '') {
    result = '';
  } else if (cellValue instanceof Date) {
    result = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy/MM/dd");
  } else if (typeof cellValue === "string") {
    // 例："4月1日" という形式の場合
    var m = cellValue.match(/^(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      var month = ("0" + m[1]).slice(-2);
      var day = ("0" + m[2]).slice(-2);
      result = year + "/" + month + "/" + day;
    } else {
      result = cellValue;
    }
  } else {
    result = cellValue.toString();
  }
  return result;
}

function findDateRow(values, targetDisplayString, year) {
  for (var i = 0; i < values.length; i++) {
    var cellString = convertCellValue(values[i][0], year);
    if (cellString === targetDisplayString) {
      return i + 1; // 行番号は1から始まる
    }
  }
  return null;
}
