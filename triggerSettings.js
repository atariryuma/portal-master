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
const DAY_NAMES = Object.freeze(['日', '月', '火', '水', '木', '金', '土']);
const MANAGED_TRIGGER_FUNCTIONS = Object.freeze(['saveToPDF', 'calculateCumulativeHours', 'syncCalendars', 'setDailyHyperlink']);

// ========================================
// メイン関数
// ========================================

/**
 * トリガー設定ダイアログを表示
 */
function showTriggerSettingsDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('triggerSettingsDialog').evaluate()
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
    const settingsSheet = getSettingsSheetOrThrow();

    // バッチ読み取り: C18:C27を一括取得（10セル）
    const values = settingsSheet.getRange('C18:C27').getValues();
    const rawSettings = {
      weeklyPdf: {
        enabled: values[0][0],  // C18
        day: values[1][0],      // C19
        hour: values[2][0]      // C20
      },
      cumulativeHours: {
        enabled: values[3][0],  // C21
        day: values[4][0],      // C22
        hour: values[5][0]      // C23
      },
      calendarSync: {
        enabled: values[6][0],  // C24
        hour: values[7][0]      // C25
      },
      dailyLink: {
        enabled: values[8][0],  // C26
        hour: values[9][0]      // C27
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
    const settingsSheet = getSettingsSheetOrThrow();

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

  hours.forEach(function(hour) {
    if (typeof hour !== 'number' || hour < 0 || hour > 23) {
      throw new Error('実行時刻は0～23の範囲で指定してください。');
    }
  });

  // 曜日の検証
  const days = [
    settings.weeklyPdf.day,
    settings.cumulativeHours.day
  ];

  days.forEach(function(day) {
    if (typeof day !== 'number' || day < 0 || day > 6) {
      throw new Error('実行曜日は0～6の範囲で指定してください。');
    }
  });
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
  // バッチ書き込み: C18:C27を一括設定（10セル）
  sheet.getRange('C18:C27').setValues([
    [settings.weeklyPdf.enabled],        // C18
    [settings.weeklyPdf.day],            // C19
    [settings.weeklyPdf.hour],           // C20
    [settings.cumulativeHours.enabled],  // C21
    [settings.cumulativeHours.day],      // C22
    [settings.cumulativeHours.hour],     // C23
    [settings.calendarSync.enabled],     // C24
    [settings.calendarSync.hour],        // C25
    [settings.dailyLink.enabled],        // C26
    [settings.dailyLink.hour]            // C27
  ]);

  Logger.log('[INFO] 設定値をシートに保存しました。');
}

// ========================================
// トリガー管理
// ========================================

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
    Logger.log('[INFO] 週報PDF保存トリガーを作成: ' + getDayName(settings.weeklyPdf.day) + '曜日 ' + settings.weeklyPdf.hour + '時');
  }

  // 2. 累計時数計算（週次）
  if (settings.cumulativeHours.enabled) {
    const trigger = ScriptApp.newTrigger('calculateCumulativeHours')
      .timeBased()
      .onWeekDay(WEEKDAY_MAP[settings.cumulativeHours.day])
      .atHour(settings.cumulativeHours.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log('[INFO] 累計時数計算トリガーを作成: ' + getDayName(settings.cumulativeHours.day) + '曜日 ' + settings.cumulativeHours.hour + '時');
  }

  // 3. カレンダー同期（毎日）
  if (settings.calendarSync.enabled) {
    const trigger = ScriptApp.newTrigger('syncCalendars')
      .timeBased()
      .everyDays(1)
      .atHour(settings.calendarSync.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log('[INFO] カレンダー同期トリガーを作成: 毎日 ' + settings.calendarSync.hour + '時');
  }

  // 4. 今日の日付へ移動（毎日）
  if (settings.dailyLink.enabled) {
    const trigger = ScriptApp.newTrigger('setDailyHyperlink')
      .timeBased()
      .everyDays(1)
      .atHour(settings.dailyLink.hour)
      .create();
    createdTriggers.push(trigger);
    Logger.log('[INFO] 今日の日付へ移動トリガーを作成: 毎日 ' + settings.dailyLink.hour + '時');
  }

  Logger.log('[INFO] 合計' + createdTriggers.length + '個のトリガーを作成しました。');
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
    rollbackCreatedTriggers_(createdTriggers);
    throw new Error('トリガーの再構築に失敗しました。既存トリガーは保持されます。詳細: ' + error.toString());
  }

  let removedCount = 0;
  const deleteErrors = [];
  existingManagedTriggers.forEach(function(trigger) {
    try {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    } catch (error) {
      deleteErrors.push(error.toString());
      Logger.log('[WARNING] 旧トリガーの削除に失敗: ' + error.toString());
    }
  });

  if (deleteErrors.length > 0) {
    const rollbackResult = rollbackCreatedTriggers_(createdTriggers);
    const details = [
      '旧トリガー削除失敗: ' + deleteErrors.length + '件',
      '新規トリガーロールバック: ' + rollbackResult.removedCount + '件'
    ];
    if (rollbackResult.errors.length > 0) {
      details.push('ロールバック失敗: ' + rollbackResult.errors.length + '件');
    }
    throw new Error('旧トリガーの削除に失敗したため、重複実行を防ぐために新規トリガーをロールバックしました。' + details.join(' / '));
  }

  return {
    createdCount: createdTriggers.length,
    removedCount: removedCount
  };
}

function rollbackCreatedTriggers_(triggers) {
  const result = {
    removedCount: 0,
    errors: []
  };

  (triggers || []).forEach(function(trigger) {
    try {
      ScriptApp.deleteTrigger(trigger);
      result.removedCount++;
    } catch (error) {
      result.errors.push(error.toString());
      Logger.log('[WARNING] 新規トリガーのロールバック削除に失敗: ' + error.toString());
    }
  });

  return result;
}

function getManagedProjectTriggers() {
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    const handler = trigger.getHandlerFunction();
    return MANAGED_TRIGGER_FUNCTIONS.indexOf(handler) !== -1;
  });
}

/**
 * 管理対象の全プロジェクトトリガーを削除
 */
function deleteAllProjectTriggers() {
  const triggers = getManagedProjectTriggers();

  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log('[INFO] 管理対象トリガーを' + triggers.length + '個削除しました。');
}

function getDefaultTriggerSettings() {
  return Object.freeze({
    weeklyPdf: Object.freeze({ enabled: true, day: 1, hour: 2 }),
    cumulativeHours: Object.freeze({ enabled: true, day: 1, hour: 2 }),
    calendarSync: Object.freeze({ enabled: true, hour: 3 }),
    dailyLink: Object.freeze({ enabled: true, hour: 4 })
  });
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
