/**
 * 共通ユーティリティ関数
 * システム全体で使用される共通関数を定義
 */

// ========================================
// 定数定義
// ========================================

/**
 * マスターシート定数
 * @const {Object}
 */
const MASTER_SHEET = Object.freeze({
  NAME: 'マスター',
  DUTY_COLUMN: 41,         // AO列 (1-based)
  DUTY_SOURCE_INDEX: 40,   // AO列 (0-based, row[40])
  INTERNAL_EVENT_INDEX: 2, // C列 (0-based, row[2])
  EXTERNAL_EVENT_INDEX: 3, // D列 (0-based, row[3])
  DATA_START_COLUMN: 5,    // E列
  DATA_COLUMN_COUNT: 36,   // E:AN = 36列
  MAX_DATA_ROW: 370,
  DATA_START_ROW: 2,
  LUNCH_INDEX: 41,         // AP列 (0-based, row[41])
  DATA_RANGE_END: 'AP'     // 全データ読み取り範囲の終端列
});

/**
 * 日直表シート定数
 * @const {Object}
 */
const DUTY_ROSTER_SHEET = Object.freeze({
  NAME: '日直表',
  NAME_COLUMN: 3,          // C列
  NUMBER_COLUMN: 4,        // D列
  OUTPUT_COLUMN: 5,        // E列
  DATA_START_ROW: 2
});

/**
 * 年間行事予定表定数
 * @const {Object}
 */
const ANNUAL_SCHEDULE = Object.freeze({
  SHEET_NAME: '年間行事予定表',
  DATA_START_ROW: 2,            // データ開始行
  DATE_COLUMN: 'B',             // B列: 日付
  DATE_INDEX: 1,                // B列 (0-based, getDataRange用)
  INTERNAL_EVENT_COLUMN: 4,     // D列: 校内行事
  EXTERNAL_EVENT_COLUMN: 13,    // M列: 対外行事
  DUTY_COLUMN: 18,              // R列: 日直
  DUTY_COLUMN_LETTER: 'R',     // R列文字表記
  ATTENDANCE_START_COLUMN: 21,  // U列: 校時データ開始
  ATTENDANCE_ROWS: 6,
  ATTENDANCE_COLS: 6,
  LUNCH_COLUMN: 27,             // AA列: 給食
  CLEAR_EVENT_RANGE: 'D',       // 年度更新クリア: 校内行事開始列
  CLEAR_EVENT_END: 'S',         // 年度更新クリア: 校内行事終了列
  CLEAR_DATA_RANGE: 'U',        // 年度更新クリア: 校時データ開始列
  CLEAR_DATA_END: 'AB',         // 年度更新クリア: 校時データ終了列
  HOLIDAY_CALENDAR_NAME: '日本の祝日'
});

/**
 * 時数様式テンプレート定数
 * @const {Object}
 */
const JISUU_TEMPLATE = Object.freeze({
  SHEET_NAME: '時数様式',
  GRADE_BLOCK_HEIGHT: 21,
  MOD_COLUMN_INDEX: 18,        // R列
  MOD_FRACTION_FORMAT: '0 ?/?',
  DATA_START_ROW: 4,
  GRADE_LABEL_ROW: 2,
  STANDARD_HOUR_ROW: 17
});

/**
 * 週報定数
 * @const {Object}
 */
const WEEKLY_REPORT = Object.freeze({
  SHEET_NAMES: Object.freeze(['週報（時数あり）', '週報（時数あり）次週']),
  TRIGGER_CELL: 'U41',
  FIRST_RANGE_START: 40,
  FIRST_RANGE_COUNT: 6,
  SECOND_RANGE_START: 57,
  SECOND_RANGE_COUNT: 6,
  MIN_HEIGHT: 6,
  MAX_HEIGHT: 14,
  NAME_RANGE: 'B1:D1',
  DATE_RANGE: 'M1:P1'
});

/**
 * 累計時数シート定数
 * @const {Object}
 */
const CUMULATIVE_SHEET = Object.freeze({
  NAME: '累計時数',
  GRADE_START_ROW: 3,
  DATE_CELL: 'A1'
});

/**
 * インポート定数
 * @const {Object}
 */
const IMPORT_CONSTANTS = Object.freeze({
  ROWS_TO_COPY: 366,
  SOURCE_SHEET_NAME: 'メインデータ'
});

/**
 * 行事カテゴリーの定義
 * @const {Object}
 */
const EVENT_CATEGORIES = Object.freeze({
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
});

/**
 * 累計対象カテゴリ（EVENT_CATEGORIESから「補習」を除外）
 * ※ トップレベルで他ファイルの定数を参照するとGASの読み込み順でエラーになるため、
 *   EVENT_CATEGORIESと同じファイルで定義する
 * @const {Array<string>}
 */
const CUMULATIVE_EVENT_CATEGORIES = Object.freeze(Object.keys(EVENT_CATEGORIES).filter(function(key) {
  return key !== '補習';
}));

/**
 * 設定シートのセル位置
 * @const {Object}
 */
const CONFIG_CELLS = Object.freeze({
  WEEKLY_REPORT_FOLDER_ID: 'C14',  // 週報フォルダID
  CALENDAR_EVENT_ID: 'C15',         // 行事予定カレンダーID
  CALENDAR_EXTERNAL_ID: 'C16'       // 対外行事カレンダーID
});

/**
 * 設定シート名
 * @const {string}
 */
const SETTINGS_SHEET_NAME = 'app_config';

/**
 * 年度更新関連設定のセル位置
 * @const {Object}
 */
const ANNUAL_UPDATE_CONFIG_CELLS = Object.freeze({
  COPY_FILE_NAME: 'C5',
  COPY_DESTINATION_FOLDER_ID: 'C7',
  BASE_SUNDAY: 'C11',
  WEEKLY_REPORT_FOLDER_ID: 'C14',
  EVENT_CALENDAR_ID: 'C15',
  EXTERNAL_CALENDAR_ID: 'C16'
});

/**
 * 自動トリガー設定のセル位置
 * @const {Object}
 */
const TRIGGER_CONFIG_CELLS = Object.freeze({
  WEEKLY_PDF_ENABLED: 'C18',
  WEEKLY_PDF_DAY: 'C19',
  WEEKLY_PDF_HOUR: 'C20',
  CUMULATIVE_HOURS_ENABLED: 'C21',
  CUMULATIVE_HOURS_DAY: 'C22',
  CUMULATIVE_HOURS_HOUR: 'C23',
  CALENDAR_SYNC_ENABLED: 'C24',
  CALENDAR_SYNC_HOUR: 'C25',
  DAILY_LINK_ENABLED: 'C26',
  DAILY_LINK_HOUR: 'C27',
  LAST_UPDATE: 'C28'
});

/**
 * 曜日番号とScriptApp.WeekDayのマッピング
 * @const {Object}
 */
const WEEKDAY_MAP = Object.freeze({
  0: ScriptApp.WeekDay.SUNDAY,
  1: ScriptApp.WeekDay.MONDAY,
  2: ScriptApp.WeekDay.TUESDAY,
  3: ScriptApp.WeekDay.WEDNESDAY,
  4: ScriptApp.WeekDay.THURSDAY,
  5: ScriptApp.WeekDay.FRIDAY,
  6: ScriptApp.WeekDay.SATURDAY
});

/**
 * 年間行事予定表のカラムインデックス（0-based）
 * @const {Object}
 */
const SCHEDULE_COLUMNS = Object.freeze({
  DATE: 0,              // 日付列（A列）
  GRADE: 19,            // 学年列（T列）
  DATA_START: 20,       // データ開始列（U列）
  DATA_END: 25          // データ終了列（Z列）
});

/**
 * モジュール学習管理の設定
 * @const {Object}
 */
const MODULE_SHEET_NAMES = Object.freeze({
  CONTROL: 'module_control',
  PLAN_SUMMARY: 'モジュール学習計画',
  // 旧シート名（移行用に保持）
  SETTINGS: 'module_settings',
  CYCLE_PLAN: 'module_cycle_plan',
  DAILY_PLAN: 'module_daily_plan',
  PLAN: 'module_plan',
  EXCEPTIONS: 'module_exceptions',
  SUMMARY: 'module_summary'
});

/**
 * module_settings シートで使用するキー
 * @const {Object}
 */
const MODULE_SETTING_KEYS = Object.freeze({
  PLAN_START_DATE: 'PLAN_START_DATE',
  PLAN_END_DATE: 'PLAN_END_DATE',
  WEEKDAYS_ENABLED: 'WEEKDAYS_ENABLED',
  LAST_GENERATED_AT: 'LAST_GENERATED_AT',
  LAST_DAILY_PLAN_COUNT: 'LAST_DAILY_PLAN_COUNT',
  DATA_VERSION: 'DATA_VERSION'
});

/**
 * モジュール学習データバージョン
 * @const {string}
 */
const MODULE_DATA_VERSION = 'CONTROL_V4';

/**
 * モジュール学習の年度開始月（4月）
 * @const {number}
 */
const MODULE_FISCAL_YEAR_START_MONTH = 4;

/**
 * 累計時数シートへのモジュール出力列（1-based）
 * @const {Object}
 */
const MODULE_CUMULATIVE_COLUMNS = Object.freeze({
  PLAN: 13,    // M列
  ACTUAL: 14,  // N列
  DIFF: 15,    // O列
  DISPLAY: 16  // P列（表示列）
});

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
    // normalizeToDate は moduleHoursDisplay.js で定義されているが、
    // GAS読み込み順が非決定的なため、ここでは自己完結型で処理する
    // ⚠ このパース処理は normalizeToDate() と同一ロジック。変更時は両方を同期すること
    let dateObj;
    if (date instanceof Date) {
      dateObj = date;
    } else if (typeof date === 'string') {
      const ymd = date.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (ymd) {
        dateObj = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
      } else {
        dateObj = new Date(date);
      }
    } else {
      dateObj = new Date(date);
    }
    if (isNaN(dateObj.getTime())) return '';
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'M月d日');
  } catch (e) {
    Logger.log('[ERROR] 日付フォーマットエラー: ' + e.toString());
    return '';
  }
}

/**
 * 次の土曜日の日付を取得（今日が土曜日の場合は今日を返す）
 * @return {Date} 土曜日の日付
 */
function getCurrentOrNextSaturday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOfWeek = today.getDay();

  if (dayOfWeek === 6) {
    // 今日が土曜日の場合
    return today;
  }

  // 次の土曜日までの日数を計算
  // (6 - dayOfWeek + 7) % 7 の説明:
  // - 日曜日(0)の場合: (6-0+7)%7 = 6日後
  // - 月曜日(1)の場合: (6-1+7)%7 = 5日後
  // - 金曜日(5)の場合: (6-5+7)%7 = 1日後
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
  return names.filter(function(name) { return name; }).join('\n');
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
  const settingsSheet = getSettingsSheetOrThrow();
  
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
  const settingsSheet = getSettingsSheetOrThrow();
  
  const cellKey = calendarType === 'EVENT' ? 
    CONFIG_CELLS.CALENDAR_EVENT_ID : 
    CONFIG_CELLS.CALENDAR_EXTERNAL_ID;
  
  const calendarName = calendarType === 'EVENT' ? 
    '行事予定カレンダー' : 
    '対外行事カレンダー';
  
  const calendarIdCell = settingsSheet.getRange(cellKey);
  let calendarId = String(calendarIdCell.getValue() || '').trim();
  
  if (calendarId) {
    const existingCalendar = CalendarApp.getCalendarById(calendarId);
    if (existingCalendar) {
      return calendarId;
    }
    Logger.log(`[WARNING] ${calendarName}IDが無効です。再検索します: ${calendarId}`);
  }

  const sameNameCalendars = CalendarApp.getCalendarsByName(calendarName);
  if (sameNameCalendars && sameNameCalendars.length > 0) {
    calendarId = sameNameCalendars[0].getId();
    calendarIdCell.setValue(calendarId);
    Logger.log(`[INFO] 既存の${calendarName}を再利用します。ID: ${calendarId}`);
    return calendarId;
  }

  Logger.log(`[INFO] ${calendarName}が見つからないため、新規作成します。`);
  const newCalendar = CalendarApp.createCalendar(calendarName);
  calendarId = newCalendar.getId();
  calendarIdCell.setValue(calendarId);
  Logger.log(`[INFO] 新規作成された${calendarName}ID: ${calendarId}`);
  
  return calendarId;
}

/**
 * 設定シート（app_config）を取得
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 設定シート
 * @throws {Error} 設定シートが見つからない場合
 */
function getSettingsSheetOrThrow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    throw new Error('設定シート（' + SETTINGS_SHEET_NAME + '）が見つかりません。');
  }

  return sheet;
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

// ========================================
// データ処理関数
// ========================================

/**
 * 日付をキーにしたマップを作成
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} dateColumn - 日付列（'A', 'B'など）
 * @return {Object} 日付をキーとしたマップ（{"M月d日": 最初の行番号}）
 */
function createDateMap(sheet, dateColumn = ANNUAL_SCHEDULE.DATE_COLUMN) {
  const lastRow = sheet.getLastRow();
  // 複数文字カラム（AA, AB等）にも対応
  let columnNumber = 0;
  for (let i = 0; i < dateColumn.length; i++) {
    columnNumber = columnNumber * 26 + (dateColumn.charCodeAt(i) - 64);
  }
  const dateValues = sheet.getRange(1, columnNumber, lastRow, 1).getValues();
  
  const dateMap = {};
  dateValues.forEach(function(row, index) {
    const date = formatDateToJapanese(row[0]);
    if (date && !Object.prototype.hasOwnProperty.call(dateMap, date)) {
      // 同一日付が複数行に存在する場合は先頭行を採用する
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
    const sheet = ss.getSheetByName(ANNUAL_SCHEDULE.SHEET_NAME);

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

/**
 * 指定されたシート名のシートを取得（エラー時は例外を投げる）
 * @param {string} sheetName - 取得するシート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 * @throws {Error} シートが見つからない場合
 */
function getSheetByNameOrThrow(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + 'シートが見つかりません。');
  }
  return sheet;
}
