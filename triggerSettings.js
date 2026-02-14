/**
 * 自動トリガー設定カスタマイズシステム
 * ユーザーが自動処理の有効/無効、実行時刻、実行曜日を設定できる機能
 */

// ========================================
// 定数定義
// ========================================

/**
 * 曜日名の配列（0=日曜、1=月曜...6=土曜）
 * @const {Array<string>}
 */
const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];
const MANAGED_TRIGGER_FUNCTIONS = ['saveToPDF', 'calculateCumulativeHours', 'syncCalendars', 'setDailyHyperlink'];

// ========================================
// メイン関数
// ========================================

/**
 * トリガー設定ダイアログを表示
 */
function showTriggerSettingsDialog() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('triggerSettingsDialog')
      .setWidth(650)
      .setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '自動トリガー設定');
  } catch (error) {
    showAlert('ダイアログの表示に失敗しました: ' + error.toString(), 'エラー');
  }
}

/**
 * 既存の設定値を取得してUIに返す
 * @return {Object} 設定値オブジェクト
 */
function getTriggerSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('年度更新作業');

    if (!settingsSheet) {
      throw new Error('年度更新作業シートが見つかりません。');
    }

    // 設定値を読み込み（空の場合はデフォルト値を使用）
    const rawSettings = {
      weeklyPdf: {
        enabled: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_ENABLED).getValue(),
        day: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_DAY).getValue(),
        hour: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_HOUR).getValue()
      },
      cumulativeHours: {
        enabled: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_ENABLED).getValue(),
        day: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_DAY).getValue(),
        hour: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_HOUR).getValue()
      },
      calendarSync: {
        enabled: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.CALENDAR_SYNC_ENABLED).getValue(),
        hour: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.CALENDAR_SYNC_HOUR).getValue()
      },
      dailyLink: {
        enabled: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.DAILY_LINK_ENABLED).getValue(),
        hour: settingsSheet.getRange(TRIGGER_CONFIG_CELLS.DAILY_LINK_HOUR).getValue()
      }
    };

    return normalizeTriggerSettings(rawSettings);

  } catch (error) {
    Logger.log('[ERROR] 設定値の取得に失敗: ' + error.toString());
    throw error;
  }
}

/**
 * トリガー設定を保存して実行
 * @param {Object} settings - 設定値オブジェクト
 * @return {string} 成功メッセージ
 */
function saveTriggerSettings(settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('年度更新作業');

    if (!settingsSheet) {
      throw new Error('年度更新作業シートが見つかりません。');
    }

    // 入力値を正規化して検証
    const normalizedSettings = normalizeTriggerSettings(settings);
    validateTriggerSettings(normalizedSettings);

    // 設定値をシートに保存
    saveTriggerSettingsToSheet(settingsSheet, normalizedSettings);

    // 安全にトリガーを再構築（新規作成成功後に既存管理トリガーを削除）
    const triggerResult = replaceManagedProjectTriggers(normalizedSettings);

    // 最終更新日時を記録
    const now = new Date();
    settingsSheet.getRange(TRIGGER_CONFIG_CELLS.LAST_UPDATE).setValue(now);

    Logger.log('[INFO] トリガー設定を更新しました: ' + now);

    return '自動トリガーの設定を保存しました。\n作成: ' + triggerResult.createdCount + '件 / 旧設定削除: ' + triggerResult.removedCount + '件';

  } catch (error) {
    Logger.log('[ERROR] トリガー設定の保存に失敗: ' + error.toString());
    throw error;
  }
}

// ========================================
// バリデーション
// ========================================

/**
 * 設定値の妥当性を検証
 * @param {Object} settings - 設定値オブジェクト
 */
function validateTriggerSettings(settings) {
  if (!settings || typeof settings !== 'object') {
    throw new Error('設定値が不正です。');
  }

  // 時刻の検証
  const hours = [
    settings.weeklyPdf.hour,
    settings.cumulativeHours.hour,
    settings.calendarSync.hour,
    settings.dailyLink.hour
  ];

  for (let hour of hours) {
    if (typeof hour !== 'number' || hour < 0 || hour > 23) {
      throw new Error('実行時刻は0～23の範囲で指定してください。');
    }
  }

  // 曜日の検証
  const days = [
    settings.weeklyPdf.day,
    settings.cumulativeHours.day
  ];

  for (let day of days) {
    if (typeof day !== 'number' || day < 0 || day > 6) {
      throw new Error('実行曜日は0～6の範囲で指定してください。');
    }
  }

  Logger.log('[INFO] 設定値の検証が完了しました。');
}

// ========================================
// 設定値の保存
// ========================================

/**
 * 設定値をシートに保存
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 設定シート
 * @param {Object} settings - 設定値オブジェクト
 */
function saveTriggerSettingsToSheet(sheet, settings) {
  // 週報PDF保存
  sheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_ENABLED).setValue(settings.weeklyPdf.enabled);
  sheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_DAY).setValue(settings.weeklyPdf.day);
  sheet.getRange(TRIGGER_CONFIG_CELLS.WEEKLY_PDF_HOUR).setValue(settings.weeklyPdf.hour);

  // 累計時数計算
  sheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_ENABLED).setValue(settings.cumulativeHours.enabled);
  sheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_DAY).setValue(settings.cumulativeHours.day);
  sheet.getRange(TRIGGER_CONFIG_CELLS.CUMULATIVE_HOURS_HOUR).setValue(settings.cumulativeHours.hour);

  // カレンダー同期
  sheet.getRange(TRIGGER_CONFIG_CELLS.CALENDAR_SYNC_ENABLED).setValue(settings.calendarSync.enabled);
  sheet.getRange(TRIGGER_CONFIG_CELLS.CALENDAR_SYNC_HOUR).setValue(settings.calendarSync.hour);

  // 今日の日付へ移動
  sheet.getRange(TRIGGER_CONFIG_CELLS.DAILY_LINK_ENABLED).setValue(settings.dailyLink.enabled);
  sheet.getRange(TRIGGER_CONFIG_CELLS.DAILY_LINK_HOUR).setValue(settings.dailyLink.hour);

  Logger.log('[INFO] 設定値をシートに保存しました。');
}

// ========================================
// トリガー管理
// ========================================

/**
 * プロジェクトの全トリガーを削除
 */
function deleteAllProjectTriggers() {
  const triggers = getManagedProjectTriggers();

  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log('[INFO] 管理対象トリガーを' + triggers.length + '個削除しました。');
}

/**
 * 設定に基づいて新しいトリガーを作成
 * @param {Object} settings - 設定値オブジェクト
 */
function createTriggersFromSettings(settings) {
  const createdTriggers = [];

  // 1. 週報PDF保存（週次）
  if (settings.weeklyPdf.enabled) {
    const trigger = ScriptApp.newTrigger('saveToPDF')
      .timeBased()
      .onWeekDay(WEEKDAY_MAP[settings.weeklyPdf.day])
      .atHour(settings.weeklyPdf.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log(`[INFO] 週報PDF保存トリガーを作成: ${getDayName(settings.weeklyPdf.day)}曜日 ${settings.weeklyPdf.hour}時`);
  }

  // 2. 累計時数計算（週次）
  if (settings.cumulativeHours.enabled) {
    const trigger = ScriptApp.newTrigger('calculateCumulativeHours')
      .timeBased()
      .onWeekDay(WEEKDAY_MAP[settings.cumulativeHours.day])
      .atHour(settings.cumulativeHours.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log(`[INFO] 累計時数計算トリガーを作成: ${getDayName(settings.cumulativeHours.day)}曜日 ${settings.cumulativeHours.hour}時`);
  }

  // 3. カレンダー同期（毎日）
  if (settings.calendarSync.enabled) {
    const trigger = ScriptApp.newTrigger('syncCalendars')
      .timeBased()
      .everyDays(1)
      .atHour(settings.calendarSync.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log(`[INFO] カレンダー同期トリガーを作成: 毎日 ${settings.calendarSync.hour}時`);
  }

  // 4. 今日の日付へ移動（毎日）
  if (settings.dailyLink.enabled) {
    const trigger = ScriptApp.newTrigger('setDailyHyperlink')
      .timeBased()
      .everyDays(1)
      .atHour(settings.dailyLink.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log(`[INFO] 今日の日付へ移動トリガーを作成: 毎日 ${settings.dailyLink.hour}時`);
  }

  Logger.log(`[INFO] 合計${createdTriggers.length}個のトリガーを作成しました。`);
  return createdTriggers;
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 曜日番号を日本語名に変換
 * @param {number} day - 曜日番号（0=日曜、1=月曜...）
 * @return {string} 曜日名
 */
function getDayName(day) {
  return DAY_NAMES[day] || '不明';
}

// ========================================
// 内部ユーティリティ
// ========================================

function replaceManagedProjectTriggers(settings) {
  const existingManagedTriggers = getManagedProjectTriggers();
  let createdTriggers = [];

  try {
    createdTriggers = createTriggersFromSettings(settings);
  } catch (error) {
    createdTriggers.forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
    throw new Error('トリガーの再構築に失敗しました。既存トリガーは保持されます。詳細: ' + error.toString());
  }

  let removedCount = 0;
  existingManagedTriggers.forEach(function(trigger) {
    try {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    } catch (error) {
      Logger.log('[WARNING] 旧トリガーの削除に失敗: ' + error.toString());
    }
  });

  return {
    createdCount: createdTriggers.length,
    removedCount: removedCount
  };
}

function getManagedProjectTriggers() {
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    const handler = trigger.getHandlerFunction();
    return MANAGED_TRIGGER_FUNCTIONS.indexOf(handler) !== -1;
  });
}

function getDefaultTriggerSettings() {
  return {
    weeklyPdf: { enabled: true, day: 1, hour: 2 },
    cumulativeHours: { enabled: true, day: 1, hour: 2 },
    calendarSync: { enabled: true, hour: 3 },
    dailyLink: { enabled: true, hour: 4 }
  };
}

function normalizeTriggerSettings(settings) {
  const defaults = getDefaultTriggerSettings();
  const input = settings || {};

  return {
    weeklyPdf: {
      enabled: toBooleanOrDefault(input.weeklyPdf && input.weeklyPdf.enabled, defaults.weeklyPdf.enabled),
      day: toIntOrDefault(input.weeklyPdf && input.weeklyPdf.day, defaults.weeklyPdf.day),
      hour: toIntOrDefault(input.weeklyPdf && input.weeklyPdf.hour, defaults.weeklyPdf.hour)
    },
    cumulativeHours: {
      enabled: toBooleanOrDefault(input.cumulativeHours && input.cumulativeHours.enabled, defaults.cumulativeHours.enabled),
      day: toIntOrDefault(input.cumulativeHours && input.cumulativeHours.day, defaults.cumulativeHours.day),
      hour: toIntOrDefault(input.cumulativeHours && input.cumulativeHours.hour, defaults.cumulativeHours.hour)
    },
    calendarSync: {
      enabled: toBooleanOrDefault(input.calendarSync && input.calendarSync.enabled, defaults.calendarSync.enabled),
      hour: toIntOrDefault(input.calendarSync && input.calendarSync.hour, defaults.calendarSync.hour)
    },
    dailyLink: {
      enabled: toBooleanOrDefault(input.dailyLink && input.dailyLink.enabled, defaults.dailyLink.enabled),
      hour: toIntOrDefault(input.dailyLink && input.dailyLink.hour, defaults.dailyLink.hour)
    }
  };
}

function toBooleanOrDefault(value, defaultValue) {
  if (value === '' || value === null || value === undefined) {
    return defaultValue;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'on') {
      return true;
    }
    if (normalized === 'false' || normalized === '0' || normalized === 'no' || normalized === 'off') {
      return false;
    }
  }

  return defaultValue;
}

function toIntOrDefault(value, defaultValue) {
  if (value === '' || value === null || value === undefined) {
    return defaultValue;
  }

  const num = Number(value);
  if (!Number.isFinite(num)) {
    return defaultValue;
  }
  return Math.floor(num);
}

// ========================================
// 後方互換性のための関数
// ========================================

/**
 * 旧形式の自動処理設定（後方互換性のため残す）
 * 実際にはダイアログを表示する
 */
function setAutomaticProcesses() {
  showTriggerSettingsDialog();
}
