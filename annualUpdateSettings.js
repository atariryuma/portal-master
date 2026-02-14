/**
 * @fileoverview 年度更新作業シートの設定値をダイアログで管理する機能
 */

/**
 * 年度更新設定ダイアログを表示
 */
function showAnnualUpdateSettingsDialog() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('annualUpdateSettingsDialog')
      .setWidth(700)
      .setHeight(760);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '年度更新設定');
  } catch (error) {
    showAlert('ダイアログの表示に失敗しました: ' + error.toString(), 'エラー');
  }
}

/**
 * 年度更新関連の設定値を取得
 * @return {Object} 設定値
 */
function getAnnualUpdateSettings() {
  try {
    const sheet = getAnnualUpdateSettingsSheet_();
    return {
      copyFileName: toTrimmedTextAnnualUpdate_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_FILE_NAME).getValue()),
      copyDestinationFolderId: toTrimmedTextAnnualUpdate_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_DESTINATION_FOLDER_ID).getValue()),
      baseSunday: formatAnnualUpdateDateForInput_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.BASE_SUNDAY).getValue()),
      weeklyReportFolderId: toTrimmedTextAnnualUpdate_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.WEEKLY_REPORT_FOLDER_ID).getValue()),
      eventCalendarId: toTrimmedTextAnnualUpdate_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.EVENT_CALENDAR_ID).getValue()),
      externalCalendarId: toTrimmedTextAnnualUpdate_(sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.EXTERNAL_CALENDAR_ID).getValue())
    };
  } catch (error) {
    Logger.log('[ERROR] 年度更新設定の取得に失敗: ' + error.toString());
    throw error;
  }
}

/**
 * 年度更新関連の設定値を保存
 * @param {Object} settings - 保存する設定値
 * @return {string} 成功メッセージ
 */
function saveAnnualUpdateSettings(settings) {
  try {
    const normalized = normalizeAnnualUpdateSettings_(settings);
    validateAnnualUpdateSettings_(normalized);

    const sheet = getAnnualUpdateSettingsSheet_();
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_FILE_NAME).setValue(normalized.copyFileName);
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.COPY_DESTINATION_FOLDER_ID).setValue(normalized.copyDestinationFolderId);
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.BASE_SUNDAY).setValue(normalized.baseSundayDate);
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.WEEKLY_REPORT_FOLDER_ID).setValue(normalized.weeklyReportFolderId);
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.EVENT_CALENDAR_ID).setValue(normalized.eventCalendarId);
    sheet.getRange(ANNUAL_UPDATE_CONFIG_CELLS.EXTERNAL_CALENDAR_ID).setValue(normalized.externalCalendarId);

    Logger.log('[INFO] 年度更新設定を保存しました。');
    return '年度更新設定を保存しました。';
  } catch (error) {
    Logger.log('[ERROR] 年度更新設定の保存に失敗: ' + error.toString());
    throw error;
  }
}

function getAnnualUpdateSettingsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('年度更新作業');
  if (!sheet) {
    throw new Error('年度更新作業シートが見つかりません。');
  }
  return sheet;
}

function normalizeAnnualUpdateSettings_(settings) {
  const input = settings || {};
  return {
    copyFileName: toTrimmedTextAnnualUpdate_(input.copyFileName),
    copyDestinationFolderId: toTrimmedTextAnnualUpdate_(input.copyDestinationFolderId),
    baseSundayDate: parseAnnualUpdateDateInput_(input.baseSunday),
    weeklyReportFolderId: toTrimmedTextAnnualUpdate_(input.weeklyReportFolderId),
    eventCalendarId: toTrimmedTextAnnualUpdate_(input.eventCalendarId),
    externalCalendarId: toTrimmedTextAnnualUpdate_(input.externalCalendarId)
  };
}

function validateAnnualUpdateSettings_(settings) {
  if (!settings.copyFileName) {
    throw new Error('複製ファイル名を入力してください。');
  }

  if (!settings.baseSundayDate) {
    throw new Error('基準日（日曜日）を入力してください。');
  }

  if (settings.baseSundayDate.getDay() !== 0) {
    throw new Error('基準日は日曜日を指定してください。');
  }

  validateFolderIdFormatIfPresent_(settings.copyDestinationFolderId, '複製先フォルダID');
  validateFolderIdFormatIfPresent_(settings.weeklyReportFolderId, '週報フォルダID');
  validateCalendarIdFormatIfPresent_(settings.eventCalendarId, '行事予定カレンダーID');
  validateCalendarIdFormatIfPresent_(settings.externalCalendarId, '対外行事カレンダーID');
}

function validateFolderIdFormatIfPresent_(folderId, fieldLabel) {
  if (!folderId) {
    return;
  }

  if (!/^[A-Za-z0-9_-]{20,}$/.test(folderId)) {
    throw new Error(fieldLabel + 'の形式が不正です。空欄にするか、正しいIDを入力してください。');
  }
}

function validateCalendarIdFormatIfPresent_(calendarId, fieldLabel) {
  if (!calendarId) {
    return;
  }

  if (calendarId.indexOf('@') === -1 || calendarId.length < 10) {
    throw new Error(fieldLabel + 'の形式が不正です。空欄にするか、正しいIDを入力してください。');
  }
}

function parseAnnualUpdateDateInput_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const date = new Date(value.getTime());
    date.setHours(0, 0, 0, 0);
    return date;
  }

  const text = toTrimmedTextAnnualUpdate_(value);
  if (!text) {
    return null;
  }

  const match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function formatAnnualUpdateDateForInput_(value) {
  const date = parseAnnualUpdateDateInput_(value);
  if (!date) {
    return '';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toTrimmedTextAnnualUpdate_(value) {
  return String(value == null ? '' : value).trim();
}
