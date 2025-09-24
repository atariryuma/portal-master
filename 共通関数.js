/**
 * 共通ユーティリティ関数
 * システム全体で使用される共通関数を定義
 */

// ========================================
// 定数定義
// ========================================

/**
 * 行事カテゴリーの定義
 * @const {Object}
 */
const EVENT_CATEGORIES = {
  "儀式": "儀式",
  "文化": "文化",
  "保健": "保健",
  "遠足": "遠足",
  "勤労": "勤労",
  "欠時数": "欠時",
  "児童会": "児童",
  "クラブ": "クラ",
  "委員会活動": "委員",
  "補習": "補習"
};

/**
 * 設定シートのセル位置
 * @const {Object}
 */
const CONFIG_CELLS = {
  WEEKLY_REPORT_FOLDER_ID: 'C14',  // 週報フォルダID
  CALENDAR_EVENT_ID: 'C15',         // 行事予定カレンダーID
  CALENDAR_EXTERNAL_ID: 'C16'       // 対外行事カレンダーID
};

// ========================================
// 日付処理関数
// ========================================

/**
 * 日付を「M月d日」形式にフォーマット
 * @param {Date|string} date - フォーマットする日付
 * @return {string} M月d日形式の文字列
 */
function formatDateToJapanese(date) {
  if (!date) return '';
  try {
    const dateObj = date instanceof Date ? date : new Date(date);
    if (isNaN(dateObj.getTime())) return '';
    return Utilities.formatDate(dateObj, 'GMT+0900', 'M月d日');
  } catch (e) {
    Logger.log('日付フォーマットエラー: ' + e.toString());
    return '';
  }
}

/**
 * 次の土曜日の日付を取得（今日が土曜日の場合は今日を返す）
 * @return {Date} 土曜日の日付
 */
function getCurrentOrNextSaturday() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  
  if (dayOfWeek === 6) {
    // 今日が土曜日の場合
    return today;
  }
  
  // 次の土曜日までの日数を計算
  const daysUntilSaturday = (6 - dayOfWeek + 7) % 7;
  const nextSaturday = new Date(today);
  nextSaturday.setDate(today.getDate() + daysUntilSaturday);
  return nextSaturday;
}

// ========================================
// 名前処理関数
// ========================================

/**
 * フルネームから名前部分を抽出
 * @param {string} fullName - フルネーム
 * @return {string} 名前部分
 */
function extractFirstName(fullName) {
  if (!fullName || typeof fullName !== 'string') return '';
  
  // 全角・半角スペースで統一的に分割
  const nameParts = fullName.trim().split(/[\s\u3000]+/);
  
  // 名前部分が存在する場合は返す
  return nameParts.length >= 2 ? nameParts[1] : '';
}

/**
 * 複数の名前を改行で結合
 * @param {Array<string>} names - 名前の配列
 * @return {string} 改行で結合された文字列
 */
function joinNamesWithNewline(names) {
  if (!Array.isArray(names)) return '';
  return names.filter(name => name).join('\n');
}

// ========================================
// フォルダ管理関数
// ========================================

/**
 * 週報フォルダIDを取得または作成
 * @return {string} フォルダID
 */
function getWeeklyReportFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('年度更新作業');
  
  if (!settingsSheet) {
    throw new Error('年度更新作業シートが見つかりません。');
  }
  
  // 設定シートからフォルダIDを取得
  const folderId = settingsSheet.getRange(CONFIG_CELLS.WEEKLY_REPORT_FOLDER_ID).getValue();
  
  if (folderId) {
    return folderId;
  }
  
  // フォルダIDが空の場合、新規作成または既存フォルダを検索
  const parentFolders = DriveApp.getFileById(ss.getId()).getParents();
  if (!parentFolders.hasNext()) {
    throw new Error('親フォルダが見つかりません。');
  }
  
  const parentFolder = parentFolders.next();
  const folderName = '週報フォルダ';
  
  // 既存フォルダを検索
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    const existingFolder = existingFolders.next();
    const existingFolderId = existingFolder.getId();
    // IDを設定シートに保存
    settingsSheet.getRange(CONFIG_CELLS.WEEKLY_REPORT_FOLDER_ID).setValue(existingFolderId);
    return existingFolderId;
  }
  
  // 新規フォルダを作成
  const newFolder = parentFolder.createFolder(folderName);
  const newFolderId = newFolder.getId();
  // IDを設定シートに保存
  settingsSheet.getRange(CONFIG_CELLS.WEEKLY_REPORT_FOLDER_ID).setValue(newFolderId);
  return newFolderId;
}

// ========================================
// カレンダー管理関数
// ========================================

/**
 * カレンダーIDを取得または作成
 * @param {string} calendarType - 'EVENT' または 'EXTERNAL'
 * @return {string} カレンダーID
 */
function getOrCreateCalendarId(calendarType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('年度更新作業');
  
  if (!settingsSheet) {
    throw new Error('年度更新作業シートが見つかりません。');
  }
  
  const cellKey = calendarType === 'EVENT' ? 
    CONFIG_CELLS.CALENDAR_EVENT_ID : 
    CONFIG_CELLS.CALENDAR_EXTERNAL_ID;
  
  const calendarName = calendarType === 'EVENT' ? 
    '行事予定カレンダー' : 
    '対外行事カレンダー';
  
  const calendarIdCell = settingsSheet.getRange(cellKey);
  let calendarId = calendarIdCell.getValue();
  
  if (!calendarId) {
    Logger.log(`[INFO] ${calendarName}が見つからないため、新規作成します。`);
    const newCalendar = CalendarApp.createCalendar(calendarName);
    calendarId = newCalendar.getId();
    calendarIdCell.setValue(calendarId);
    Logger.log(`[INFO] 新規作成された${calendarName}ID: ${calendarId}`);
  }
  
  return calendarId;
}

// ========================================
// エラーハンドリング関数
// ========================================

/**
 * UIアラートを表示（UIが利用できない場合はログ出力）
 * @param {string} message - 表示するメッセージ
 * @param {string} title - アラートのタイトル（省略可能）
 */
function showAlert(message, title = '通知') {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // UIが利用できない場合（トリガー実行時など）はログに出力
    Logger.log(`[${title}] ${message}`);
  }
}

/**
 * エラーを安全に処理してログに記録
 * @param {Function} func - 実行する関数
 * @param {string} functionName - 関数名（ログ用）
 * @return {*} 関数の実行結果
 */
function safeExecute(func, functionName) {
  try {
    return func();
  } catch (error) {
    const errorMessage = `${functionName}でエラーが発生しました: ${error.toString()}`;
    Logger.log(`[ERROR] ${errorMessage}`);
    showAlert(errorMessage, 'エラー');
    throw error;
  }
}

// ========================================
// データ処理関数
// ========================================

/**
 * 日付をキーにしたマップを作成
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} dateColumn - 日付列（'A', 'B'など）
 * @return {Object} 日付をキーとしたマップ（{"M月d日": 行番号}）
 */
function createDateMap(sheet, dateColumn = 'B') {
  const lastRow = sheet.getLastRow();
  const columnNumber = dateColumn.charCodeAt(0) - 64; // 'A' = 1, 'B' = 2
  const dateValues = sheet.getRange(1, columnNumber, lastRow, 1).getValues();
  
  const dateMap = {};
  dateValues.forEach((row, index) => {
    const date = formatDateToJapanese(row[0]);
    if (date) {
      dateMap[date] = index + 1; // 1-based index
    }
  });
  
  return dateMap;
}

/**
 * 全角文字を半角文字に変換
 * @param {string} str - 変換する文字列
 * @return {string} 変換後の文字列
 */
function convertFullWidthToHalfWidth(str) {
  if (!str) return '';

  return str.replace(/[！-～]/g, function(tmpStr) {
    return String.fromCharCode(tmpStr.charCodeAt(0) - 0xFEE0);
  })
  .replace(/"/g, '"')
  .replace(/'/g, "'")
  .replace(/'/g, "`")
  .replace(/￥/g, "\\")
  .replace(/　/g, " ")
  .replace(/〜/g, "~");
}

// ========================================
// シート管理関数
// ========================================

/**
 * 年間行事予定表シートを安全に取得
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} シートオブジェクトまたはnull
 */
function getAnnualScheduleSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('年間行事予定表');

    if (!sheet) {
      Logger.log('[WARNING] 年間行事予定表シートが見つかりません');
      return null;
    }

    // シートの基本的な妥当性をチェック
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 10) {
      Logger.log('[WARNING] 年間行事予定表シートのデータが不完全です（行数: ' + lastRow + ', 列数: ' + lastColumn + '）');
      return null;
    }

    Logger.log('[INFO] 年間行事予定表シートを正常に取得しました（行数: ' + lastRow + ', 列数: ' + lastColumn + '）');
    return sheet;

  } catch (error) {
    Logger.log('[ERROR] 年間行事予定表シート取得エラー: ' + error.toString());
    return null;
  }
}

/**
 * 年間行事予定表シートを安全に取得（エラー時は例外を投げる）
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 * @throws {Error} シートが見つからない場合
 */
function getAnnualScheduleSheetOrThrow() {
  const sheet = getAnnualScheduleSheet();
  if (!sheet) {
    throw new Error('年間行事予定表シートが見つからないか、データが不完全です。シートの存在とデータの妥当性を確認してください。');
  }
  return sheet;
}