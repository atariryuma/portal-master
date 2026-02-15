/**
 * ポータルマスター 包括的テストスイート
 * すべての機能が正常に動作しているかを確認
 */

// ========================================
// テスト実行メイン関数
// ========================================

/**
 * すべてのテストを実行
 * メニューから実行: テスト → 全機能テスト実行
 */
function runAllTests() {
  Logger.clear();
  Logger.log('====================================');
  Logger.log('ポータルマスター 全機能テスト開始');
  Logger.log('実行日時: ' + new Date());
  Logger.log('====================================\n');

  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    skipped: 0,
    errors: []
  };

  try {
    runTestGroups_(results, getFullTestPlan_());

  } catch (error) {
    Logger.log('❌ テスト実行中に致命的エラー: ' + error.toString());
    results.errors.push('致命的エラー: ' + error.toString());
  }

  // 最終結果サマリー
  Logger.log('\n====================================');
  Logger.log('テスト結果サマリー');
  Logger.log('====================================');
  Logger.log('総テスト数: ' + results.total);
  Logger.log('✅ 成功: ' + results.passed);
  Logger.log('❌ 失敗: ' + results.failed);
  Logger.log('⏭️  スキップ: ' + results.skipped);

  if (results.errors.length > 0) {
    Logger.log('\n【エラー詳細】');
    results.errors.forEach(function(error, index) {
      Logger.log((index + 1) + '. ' + error);
    });
  }

  const successRate = results.total > 0 ? Math.round((results.passed / results.total) * 100) : 0;
  Logger.log('\n成功率: ' + successRate + '%');

  if (results.failed === 0) {
    Logger.log('\n🎉 すべてのテストが成功しました！');
  } else {
    Logger.log('\n⚠️  一部のテストが失敗しています。上記のエラー詳細を確認してください。');
  }

  Logger.log('====================================\n');

  // UIにも結果を表示
  const ui = SpreadsheetApp.getUi();
  const message = 'テスト完了\n\n' +
                  '総テスト数: ' + results.total + '\n' +
                  '✅ 成功: ' + results.passed + '\n' +
                  '❌ 失敗: ' + results.failed + '\n' +
                  '成功率: ' + successRate + '%\n\n' +
                  '詳細はスクリプトエディタのログを確認してください。';

  if (results.failed === 0) {
    ui.alert('✅ テスト成功', message, ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ テスト失敗あり', message, ui.ButtonSet.OK);
  }
}

function runTestGroups_(results, groups) {
  groups.forEach(function(group) {
    Logger.log(group.title);
    group.tests.forEach(function(testItem) {
      runTest(results, testItem.name, testItem.fn);
    });
    Logger.log('');
  });
}

function getFullTestPlan_() {
  return [
    {
      title: '【フェーズ1】環境チェック',
      tests: [
        { name: '1-1. スプレッドシート取得', fn: testGetSpreadsheet },
        { name: '1-2. 必須シート存在確認', fn: testRequiredSheets },
        { name: '1-3. 設定シート構造確認', fn: testConfigSheetStructure }
      ]
    },
    {
      title: '【フェーズ2】モジュール時数統合検証',
      tests: [
        { name: '2-1. モジュール定数整合性', fn: testModuleConstants },
        { name: '2-2. モジュールシート初期化確認', fn: testInitializeModuleSheets },
        { name: '2-3. 累計時数へのMOD統合確認', fn: testModuleCumulativeIntegration },
        { name: '2-4. 表示フォーマット関数確認', fn: testModuleDisplayFormatter },
        { name: '2-5. 45分換算関数確認', fn: testSessionsToUnits },
        { name: '2-6. 表示列の占有衝突回避', fn: testResolveDisplayColumnSkipsOccupiedColumn },
        { name: '2-7. 旧スキーマ行の置換除外', fn: testReplaceRowsDropsLegacyFiscalRows },
        { name: '2-8. 他年度旧スキーマ行の保持', fn: testReplaceRowsKeepsOtherLegacyFiscalRows }
      ]
    },
    {
      title: '【フェーズ3】学年別集計・データ処理',
      tests: [
        { name: '3-1. 年間行事予定表シート取得', fn: testGetAnnualScheduleSheet },
        { name: '3-2. 日付マップ作成', fn: testCreateDateMap },
        { name: '3-3. 重複日付の先頭行マッピング', fn: testCreateDateMapKeepsFirstRow },
        { name: '3-4. イベントカテゴリ定数確認', fn: testEventCategories },
        { name: '3-5. 集計期間バリデーション（不正日付）', fn: testValidateAggregateDateRangeRejectsInvalidDate },
        { name: '3-6. 集計期間バリデーション（日付順）', fn: testValidateAggregateDateRangeRejectsReverseRange },
        { name: '3-7. 集計期間バリデーション（正常系）', fn: testValidateAggregateDateRangeAcceptsValidRange },
        { name: '3-8. 月キー生成（年度跨ぎ）', fn: testBuildMonthKeysForAggregateAcrossFiscalYear },
        { name: '3-9. 月キー生成（単月）', fn: testBuildMonthKeysForAggregateSingleMonth },
        { name: '3-10. 既存MOD値の月別退避', fn: testCaptureExistingModValuesByMonth },
        { name: '3-11. MOD実績取得関数', fn: testGetModuleActualUnitsForMonth }
      ]
    },
    {
      title: '【フェーズ4】設定・バリデーション',
      tests: [
        { name: '4-1. トリガー設定定数の存在確認', fn: testTriggerConfigConstants },
        { name: '4-2. トリガー設定値読み込み', fn: testGetTriggerSettings },
        { name: '4-3. トリガー設定バリデーション', fn: testValidateTriggerSettings },
        { name: '4-4. トリガー設定正規化', fn: testNormalizeTriggerSettings },
        { name: '4-5. 年度更新設定定数の存在確認', fn: testAnnualUpdateConfigConstants },
        { name: '4-6. 年度更新設定バリデーション', fn: testValidateAnnualUpdateSettings }
      ]
    },
    {
      title: '【フェーズ5】共通関数',
      tests: [
        { name: '5-1. 日付フォーマット関数', fn: testFormatDateToJapanese },
        { name: '5-2. 名前抽出関数', fn: testExtractFirstName },
        { name: '5-3. アラート関数定義確認', fn: testShowAlert }
      ]
    }
  ];
}

function getQuickTestPlan_() {
  return [
    {
      title: '【クイック】環境',
      tests: [
        { name: 'Q-1. スプレッドシート取得', fn: testGetSpreadsheet },
        { name: 'Q-2. 必須シート存在確認', fn: testRequiredSheets }
      ]
    },
    {
      title: '【クイック】主要ロジック',
      tests: [
        { name: 'Q-3. 累計時数へのMOD統合確認', fn: testModuleCumulativeIntegration },
        { name: 'Q-4. 集計期間バリデーション（不正日付）', fn: testValidateAggregateDateRangeRejectsInvalidDate },
        { name: 'Q-5. 集計期間バリデーション（日付順）', fn: testValidateAggregateDateRangeRejectsReverseRange },
        { name: 'Q-6. 既存MOD値の月別退避', fn: testCaptureExistingModValuesByMonth }
      ]
    }
  ];
}

/**
 * 個別テストを実行してログ出力
 */
function runTest(results, testName, testFunction) {
  results.total++;

  try {
    const result = testFunction();

    if (result.skip) {
      Logger.log('⏭️  SKIP: ' + testName + ' - ' + result.message);
      results.skipped++;
    } else if (result.success) {
      Logger.log('✅ PASS: ' + testName + (result.message ? ' - ' + result.message : ''));
      results.passed++;
    } else {
      Logger.log('❌ FAIL: ' + testName + ' - ' + result.message);
      results.failed++;
      results.errors.push(testName + ': ' + result.message);
    }
  } catch (error) {
    Logger.log('❌ ERROR: ' + testName + ' - ' + error.toString());
    results.failed++;
    results.errors.push(testName + ': ' + error.toString());
  }
}

// ========================================
// フェーズ1: 環境チェック
// ========================================

function testGetSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return { success: false, message: 'スプレッドシートが取得できません' };
  }
  return { success: true, message: 'ID: ' + ss.getId() };
}

function testRequiredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['マスター', '年度更新作業', '時数様式'];
  const missingSheets = [];

  requiredSheets.forEach(function(sheetName) {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  });

  if (missingSheets.length > 0) {
    return { success: false, message: '不足シート: ' + missingSheets.join(', ') };
  }

  return { success: true, message: requiredSheets.length + '個の必須シートを確認' };
}

function testConfigSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('年度更新作業');

  if (!sheet) {
    return { success: false, message: '年度更新作業シートが見つかりません' };
  }

  // 年度更新設定セル + トリガー設定セルの確認
  const cells = [
    'C5', 'C7', 'C11', 'C14', 'C15', 'C16',
    'C18', 'C19', 'C20', 'C21', 'C22', 'C23', 'C24', 'C25', 'C26', 'C27'
  ];
  const accessible = cells.every(function(cell) {
    try {
      sheet.getRange(cell);
      return true;
    } catch (e) {
      return false;
    }
  });

  if (!accessible) {
    return { success: false, message: '設定セルにアクセスできません' };
  }

  return { success: true, message: cells.length + '個の設定セルを確認' };
}

// ========================================
// フェーズ2: モジュール時数統合検証
// ========================================

function testModuleConstants() {
  const requiredConstants = [
    'MODULE_SHEET_NAMES',
    'MODULE_SETTING_KEYS',
    'MODULE_DATA_VERSION',
    'MODULE_FISCAL_YEAR_START_MONTH',
    'MODULE_CUMULATIVE_COLUMNS'
  ];

  const missingConstants = requiredConstants.filter(function(constantName) {
    return typeof eval(constantName) === 'undefined';
  });

  if (missingConstants.length > 0) {
    return { success: false, message: '不足定数: ' + missingConstants.join(', ') };
  }

  if (MODULE_FISCAL_YEAR_START_MONTH !== 4) {
    return { success: false, message: '年度開始月が4月固定になっていません' };
  }

  return { success: true, message: requiredConstants.length + '個のモジュール定数を確認' };
}

function testInitializeModuleSheets() {
  if (typeof initializeModuleHoursSheetsIfNeeded !== 'function') {
    return { success: false, message: 'initializeModuleHoursSheetsIfNeeded関数が見つかりません' };
  }

  try {
    initializeModuleHoursSheetsIfNeeded();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = [
      MODULE_SHEET_NAMES.CONTROL
    ];

    const missingSheets = requiredSheets.filter(function(sheetName) {
      return !ss.getSheetByName(sheetName);
    });

    if (missingSheets.length > 0) {
      return { success: false, message: '作成失敗シート: ' + missingSheets.join(', ') };
    }

    return { success: true, message: 'module_control シートを確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testModuleCumulativeIntegration() {
  if (typeof syncModuleHoursWithCumulative !== 'function') {
    return { success: false, message: 'syncModuleHoursWithCumulative関数が見つかりません' };
  }

  try {
    syncModuleHoursWithCumulative(new Date());
    const cumulativeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('累計時数');
    if (!cumulativeSheet) {
      return { success: false, message: '累計時数シートが見つかりません' };
    }

    const headers = cumulativeSheet.getRange(2, MODULE_CUMULATIVE_COLUMNS.PLAN, 1, 3).getValues()[0];
    const expectedHeaders = ['MOD計画累計', 'MOD実施累計', 'MOD差分'];
    const mismatch = expectedHeaders.filter(function(header, index) {
      return headers[index] !== header;
    });

    if (mismatch.length > 0) {
      return { success: false, message: '累計時数シートのMOD列ヘッダーが不正です' };
    }

    const displayHeaderRow = cumulativeSheet.getRange(2, 1, 1, cumulativeSheet.getLastColumn()).getValues()[0];
    if (displayHeaderRow.indexOf('MOD実施累計(表示)') === -1) {
      return { success: false, message: 'MOD実施累計(表示)列が作成されていません' };
    }

    return { success: true, message: '累計時数シートへMOD列を統合' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testModuleDisplayFormatter() {
  if (typeof formatSessionsAsMixedFraction !== 'function') {
    return { success: false, message: 'formatSessionsAsMixedFraction関数が見つかりません' };
  }

  const case1 = formatSessionsAsMixedFraction(56); // 56/3 = 18 2/3
  const case2 = formatSessionsAsMixedFraction(1);  // 1/3

  if (case1 !== '18 2/3') {
    return { success: false, message: '56セッションの表示が不正です: ' + case1 };
  }
  if (case2 !== '1/3') {
    return { success: false, message: '1セッションの表示が不正です: ' + case2 };
  }

  return { success: true, message: '表示フォーマットを確認' };
}

function testSessionsToUnits() {
  if (typeof sessionsToUnits !== 'function') {
    return { success: false, message: 'sessionsToUnits関数が見つかりません' };
  }

  const value1 = sessionsToUnits(3);    // 1
  const value2 = sessionsToUnits(1);    // 0.333...
  const value3 = sessionsToUnits('6');  // 2

  if (value1 !== 1) {
    return { success: false, message: '3セッション換算が不正です: ' + value1 };
  }
  if (Math.abs(value2 - 0.333333) > 0.000001) {
    return { success: false, message: '1セッション換算が不正です: ' + value2 };
  }
  if (value3 !== 2) {
    return { success: false, message: '文字列入力換算が不正です: ' + value3 };
  }

  return { success: true, message: '45分換算ロジックを確認' };
}

function testResolveDisplayColumnSkipsOccupiedColumn() {
  if (typeof resolveCumulativeDisplayColumn !== 'function' ||
      typeof upsertModuleSettingsValues !== 'function' ||
      typeof readModuleSettingsMap !== 'function' ||
      typeof initializeModuleHoursSheetsIfNeeded !== 'function') {
    return { success: false, message: '必要関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('tmp_mod_display_col_' + Date.now());
  const sheets = initializeModuleHoursSheetsIfNeeded();
  const settingsSheet = sheets.settingsSheet;
  const settingsMap = readModuleSettingsMap(settingsSheet);
  const previous = Object.prototype.hasOwnProperty.call(settingsMap, MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN)
    ? settingsMap[MODULE_SETTING_KEYS.CUMULATIVE_DISPLAY_COLUMN]
    : '';

  try {
    const occupiedColumn = MODULE_CUMULATIVE_COLUMNS.DISPLAY_FALLBACK;
    tempSheet.getRange(3, occupiedColumn).setValue('既存データ');
    upsertModuleSettingsValues(settingsSheet, {
      CUMULATIVE_DISPLAY_COLUMN: occupiedColumn
    });

    const resolved = resolveCumulativeDisplayColumn(tempSheet);
    if (resolved === occupiedColumn) {
      return { success: false, message: 'データ占有列を再利用しています（列: ' + resolved + '）' };
    }

    const resolvedHeader = tempSheet.getRange(2, resolved).getValue();
    if (resolvedHeader !== 'MOD実施累計(表示)') {
      return { success: false, message: '解決列のヘッダー設定が不正です: ' + resolvedHeader };
    }

    return { success: true, message: '占有列を避けて表示列を解決' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    upsertModuleSettingsValues(settingsSheet, {
      CUMULATIVE_DISPLAY_COLUMN: previous
    });
    ss.deleteSheet(tempSheet);
  }
}

function testReplaceRowsDropsLegacyFiscalRows() {
  if (typeof replaceRowsForFiscalYear !== 'function') {
    return { success: false, message: 'replaceRowsForFiscalYear関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('tmp_mod_replace_' + Date.now());

  try {
    tempSheet.getRange(1, 1, 1, 2).setValues([['fiscal_year', 'value']]);
    tempSheet.getRange(2, 1, 3, 2).setValues([
      ['2025-06', 'legacy'],
      [2024, 'keep'],
      [2025, 'old']
    ]);

    replaceRowsForFiscalYear(tempSheet, [[2025, 'new']], 2025, 0, 2);

    const afterLastRow = tempSheet.getLastRow();
    const values = afterLastRow > 1 ? tempSheet.getRange(2, 1, afterLastRow - 1, 2).getValues() : [];
    const legacyExists = values.some(function(row) {
      return String(row[0]) === '2025-06';
    });
    const oldTargetExists = values.some(function(row) {
      return Number(row[0]) === 2025 && row[1] === 'old';
    });
    const keepExists = values.some(function(row) {
      return Number(row[0]) === 2024 && row[1] === 'keep';
    });
    const newExists = values.some(function(row) {
      return Number(row[0]) === 2025 && row[1] === 'new';
    });

    if (legacyExists || oldTargetExists || !keepExists || !newExists) {
      return { success: false, message: '置換結果が不正です: ' + JSON.stringify(values) };
    }

    return { success: true, message: '旧スキーマ行を除外して年度置換できることを確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    ss.deleteSheet(tempSheet);
  }
}

function testReplaceRowsKeepsOtherLegacyFiscalRows() {
  if (typeof replaceRowsForFiscalYear !== 'function') {
    return { success: false, message: 'replaceRowsForFiscalYear関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('tmp_mod_replace_keep_' + Date.now());

  try {
    tempSheet.getRange(1, 1, 1, 2).setValues([['fiscal_year', 'value']]);
    tempSheet.getRange(2, 1, 3, 2).setValues([
      ['2024-12', 'legacy_keep'],
      ['unknown', 'opaque_keep'],
      [2025, 'old_target']
    ]);

    replaceRowsForFiscalYear(tempSheet, [[2025, 'new_target']], 2025, 0, 2);

    const afterLastRow = tempSheet.getLastRow();
    const values = afterLastRow > 1 ? tempSheet.getRange(2, 1, afterLastRow - 1, 2).getValues() : [];
    const isLegacyMonthValue = function(value, year, month) {
      if (value instanceof Date) {
        return value.getFullYear() === year && (value.getMonth() + 1) === month;
      }
      const text = String(value === null || value === undefined ? '' : value).trim();
      if (!text) {
        return false;
      }
      return text.indexOf(year + '-' + String(month).padStart(2, '0')) === 0 ||
        text.indexOf(year + '/' + month) === 0 ||
        text.indexOf(year + '/' + String(month).padStart(2, '0')) === 0;
    };
    const legacyKeepExists = values.some(function(row) {
      return isLegacyMonthValue(row[0], 2024, 12) && row[1] === 'legacy_keep';
    });
    const opaqueKeepExists = values.some(function(row) {
      return String(row[0]) === 'unknown' && row[1] === 'opaque_keep';
    });
    const oldTargetExists = values.some(function(row) {
      return Number(row[0]) === 2025 && row[1] === 'old_target';
    });
    const newTargetExists = values.some(function(row) {
      return Number(row[0]) === 2025 && row[1] === 'new_target';
    });

    if (!legacyKeepExists || !opaqueKeepExists || oldTargetExists || !newTargetExists) {
      return { success: false, message: '保持/置換結果が不正です: ' + JSON.stringify(values) };
    }

    return { success: true, message: '対象年度のみ置換し、他年度旧スキーマ行を保持' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    ss.deleteSheet(tempSheet);
  }
}

// ========================================
// フェーズ3: 新機能テスト
// ========================================

function testTriggerConfigConstants() {
  if (typeof TRIGGER_CONFIG_CELLS === 'undefined') {
    return { success: false, message: 'TRIGGER_CONFIG_CELLS定数が見つかりません' };
  }

  if (typeof WEEKDAY_MAP === 'undefined') {
    return { success: false, message: 'WEEKDAY_MAP定数が見つかりません' };
  }

  const requiredKeys = ['WEEKLY_PDF_ENABLED', 'WEEKLY_PDF_DAY', 'WEEKLY_PDF_HOUR',
                       'CUMULATIVE_HOURS_ENABLED', 'CUMULATIVE_HOURS_DAY', 'CUMULATIVE_HOURS_HOUR',
                       'CALENDAR_SYNC_ENABLED', 'CALENDAR_SYNC_HOUR',
                       'DAILY_LINK_ENABLED', 'DAILY_LINK_HOUR', 'LAST_UPDATE'];

  const missingKeys = requiredKeys.filter(function(key) {
    return !TRIGGER_CONFIG_CELLS.hasOwnProperty(key);
  });

  if (missingKeys.length > 0) {
    return { success: false, message: '不足キー: ' + missingKeys.join(', ') };
  }

  return { success: true, message: requiredKeys.length + '個の設定キーを確認' };
}

function testGetTriggerSettings() {
  try {
    const settings = getTriggerSettings();

    if (!settings || typeof settings !== 'object') {
      return { success: false, message: '設定オブジェクトが取得できません' };
    }

    const requiredSections = ['weeklyPdf', 'cumulativeHours', 'calendarSync', 'dailyLink'];
    const missingSections = requiredSections.filter(function(section) {
      return !settings.hasOwnProperty(section);
    });

    if (missingSections.length > 0) {
      return { success: false, message: '不足セクション: ' + missingSections.join(', ') };
    }

    return { success: true, message: '設定値を正常に取得' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testValidateTriggerSettings() {
  try {
    // 正常な設定値
    const validSettings = {
      weeklyPdf: { enabled: true, day: 1, hour: 2 },
      cumulativeHours: { enabled: true, day: 1, hour: 2 },
      calendarSync: { enabled: true, hour: 3 },
      dailyLink: { enabled: true, hour: 4 }
    };

    validateTriggerSettings(validSettings);

    // 異常な設定値（時刻が不正）
    const invalidSettings = {
      weeklyPdf: { enabled: true, day: 1, hour: 25 }, // 25時は存在しない
      cumulativeHours: { enabled: true, day: 1, hour: 2 },
      calendarSync: { enabled: true, hour: 3 },
      dailyLink: { enabled: true, hour: 4 }
    };

    try {
      validateTriggerSettings(invalidSettings);
      return { success: false, message: '不正な設定値を検出できませんでした' };
    } catch (validationError) {
      // エラーが投げられれば正常
    }

    return { success: true, message: 'バリデーションが正常に動作' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testNormalizeTriggerSettings() {
  if (typeof normalizeTriggerSettings !== 'function') {
    return { success: false, message: 'normalizeTriggerSettings関数が見つかりません' };
  }

  const normalized = normalizeTriggerSettings({
    weeklyPdf: { enabled: 'false', day: '2', hour: '7.9' },
    cumulativeHours: { enabled: '1', day: '', hour: '' },
    calendarSync: { enabled: 0, hour: '22' },
    dailyLink: {}
  });

  if (normalized.weeklyPdf.enabled !== false || normalized.weeklyPdf.day !== 2 || normalized.weeklyPdf.hour !== 7) {
    return { success: false, message: 'weeklyPdfの正規化が不正です: ' + JSON.stringify(normalized.weeklyPdf) };
  }
  if (normalized.cumulativeHours.enabled !== true || normalized.cumulativeHours.day !== 1 || normalized.cumulativeHours.hour !== 2) {
    return { success: false, message: 'cumulativeHoursの正規化が不正です: ' + JSON.stringify(normalized.cumulativeHours) };
  }
  if (normalized.calendarSync.enabled !== false || normalized.calendarSync.hour !== 22) {
    return { success: false, message: 'calendarSyncの正規化が不正です: ' + JSON.stringify(normalized.calendarSync) };
  }
  if (normalized.dailyLink.enabled !== true || normalized.dailyLink.hour !== 4) {
    return { success: false, message: 'dailyLinkのデフォルト補完が不正です: ' + JSON.stringify(normalized.dailyLink) };
  }

  return { success: true, message: 'トリガー設定の正規化を確認' };
}

function testAnnualUpdateConfigConstants() {
  if (typeof ANNUAL_UPDATE_CONFIG_CELLS === 'undefined') {
    return { success: false, message: 'ANNUAL_UPDATE_CONFIG_CELLS定数が見つかりません' };
  }

  const requiredKeys = [
    'COPY_FILE_NAME',
    'COPY_DESTINATION_FOLDER_ID',
    'BASE_SUNDAY',
    'WEEKLY_REPORT_FOLDER_ID',
    'EVENT_CALENDAR_ID',
    'EXTERNAL_CALENDAR_ID'
  ];

  const missingKeys = requiredKeys.filter(function(key) {
    return !ANNUAL_UPDATE_CONFIG_CELLS.hasOwnProperty(key);
  });

  if (missingKeys.length > 0) {
    return { success: false, message: '不足キー: ' + missingKeys.join(', ') };
  }

  return { success: true, message: requiredKeys.length + '個の年度更新設定キーを確認' };
}

function testValidateAnnualUpdateSettings() {
  if (typeof validateAnnualUpdateSettings_ !== 'function') {
    return { success: false, message: 'validateAnnualUpdateSettings_関数が見つかりません' };
  }

  const validSunday = new Date(2026, 1, 15); // 2026-02-15 (日)
  const invalidMonday = new Date(2026, 1, 16); // 2026-02-16 (月)

  try {
    validateAnnualUpdateSettings_({
      copyFileName: 'テスト',
      baseSundayDate: validSunday,
      copyDestinationFolderId: '',
      weeklyReportFolderId: '',
      eventCalendarId: '',
      externalCalendarId: ''
    });
  } catch (error) {
    return { success: false, message: '正常値で例外が発生しました: ' + error.toString() };
  }

  try {
    validateAnnualUpdateSettings_({
      copyFileName: 'テスト',
      baseSundayDate: invalidMonday,
      copyDestinationFolderId: '',
      weeklyReportFolderId: '',
      eventCalendarId: '',
      externalCalendarId: ''
    });
    return { success: false, message: '非日曜日を検出できませんでした' };
  } catch (error) {
    const message = error && error.message ? error.message : String(error || '');
    if (message.indexOf('基準日は日曜日を指定してください。') === -1) {
      return { success: false, message: '期待外のエラーメッセージ: ' + message };
    }
  }

  return { success: true, message: '年度更新設定の日曜日制約を確認' };
}

// ========================================
// フェーズ4: 共通関数テスト
// ========================================

function testFormatDateToJapanese() {
  if (typeof formatDateToJapanese !== 'function') {
    return { success: false, message: '関数が見つかりません' };
  }

  const testDate = new Date(2025, 0, 18); // 2025年1月18日
  const formatted = formatDateToJapanese(testDate);

  // 実装は「M月d日」形式を返す
  if (formatted !== '1月18日') {
    return { success: false, message: '期待値: 1月18日, 実際: ' + formatted };
  }

  return { success: true, message: '日付フォーマット正常（M月d日形式）' };
}

function testExtractFirstName() {
  if (typeof extractFirstName !== 'function') {
    return { success: false, message: '関数が見つかりません' };
  }

  // スペース区切りのみ対応（実装の仕様）
  const testCases = [
    { input: '山田　太郎', expected: '太郎' },  // 全角スペース
    { input: '山田 太郎', expected: '太郎' },   // 半角スペース
    { input: '佐藤　花子', expected: '花子' }  // 全角スペース
  ];

  for (var i = 0; i < testCases.length; i++) {
    const result = extractFirstName(testCases[i].input);
    if (result !== testCases[i].expected) {
      return { success: false, message: '入力: ' + testCases[i].input + ', 期待値: ' + testCases[i].expected + ', 実際: ' + result };
    }
  }

  return { success: true, message: testCases.length + '件のテストケースが成功' };
}

function testShowAlert() {
  if (typeof showAlert !== 'function') {
    return { success: false, message: '関数が見つかりません' };
  }

  return { success: true, message: '関数が定義されています' };
}

// ========================================
// フェーズ5: データ処理テスト
// ========================================

function testGetAnnualScheduleSheet() {
  if (typeof getAnnualScheduleSheet !== 'function') {
    return { success: false, message: '関数が見つかりません' };
  }

  const sheet = getAnnualScheduleSheet();
  if (!sheet) {
    return { success: false, message: '年間行事予定表シートを取得できません' };
  }

  return { success: true, message: 'シート名: ' + sheet.getName() };
}

function testCreateDateMap() {
  if (typeof createDateMap !== 'function') {
    return { success: false, message: '関数が見つかりません' };
  }

  try {
    const sheet = getAnnualScheduleSheet();
    if (!sheet) {
      return { skip: true, message: '年間行事予定表シートが見つかりません' };
    }

    const dateMap = createDateMap(sheet, 'B');

    if (!dateMap || typeof dateMap !== 'object') {
      return { success: false, message: '日付マップが作成できません' };
    }

    const dateCount = Object.keys(dateMap).length;
    return { success: true, message: dateCount + '件の日付をマッピング' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testCreateDateMapKeepsFirstRow() {
  if (typeof createDateMap !== 'function' || typeof formatDateToJapanese !== 'function') {
    return { success: false, message: '必要関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('tmp_date_map_test_' + Date.now());

  try {
    const firstDate = new Date(2025, 3, 1);
    const secondDate = new Date(2025, 3, 2);
    tempSheet.getRange(1, 2, 3, 1).setValues([[firstDate], [firstDate], [secondDate]]);

    const dateMap = createDateMap(tempSheet, 'B');
    const firstKey = formatDateToJapanese(firstDate);
    const secondKey = formatDateToJapanese(secondDate);

    if (dateMap[firstKey] !== 1) {
      return { success: false, message: '重複日付の先頭行を参照していません（期待:1, 実際:' + dateMap[firstKey] + '）' };
    }
    if (dateMap[secondKey] !== 3) {
      return { success: false, message: '2件目の日付マッピングが不正です（期待:3, 実際:' + dateMap[secondKey] + '）' };
    }

    return { success: true, message: '重複日付は先頭行に正しくマッピングされます' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    ss.deleteSheet(tempSheet);
  }
}

function testEventCategories() {
  if (typeof EVENT_CATEGORIES === 'undefined') {
    return { success: false, message: 'EVENT_CATEGORIES定数が見つかりません' };
  }

  const requiredCategories = ['儀式', '文化', '保健', '遠足', '勤労', '欠時数', '児童会', 'クラブ', '委員会活動', '補習'];
  const missingCategories = requiredCategories.filter(function(cat) {
    return !EVENT_CATEGORIES.hasOwnProperty(cat);
  });

  if (missingCategories.length > 0) {
    return { success: false, message: '不足カテゴリ: ' + missingCategories.join(', ') };
  }

  return { success: true, message: requiredCategories.length + '個のカテゴリを確認' };
}

function testValidateAggregateDateRangeRejectsInvalidDate() {
  if (typeof parseAndValidateAggregateDateRange !== 'function') {
    return { success: false, message: 'parseAndValidateAggregateDateRange関数が見つかりません' };
  }

  try {
    parseAndValidateAggregateDateRange('invalid-date', '2026-03-31');
    return { success: false, message: '不正日付を検出できませんでした' };
  } catch (error) {
    const message = error && error.message ? error.message : String(error || '');
    if (message.indexOf('入力された日付が無効です。') === -1) {
      return { success: false, message: '期待外のエラーメッセージ: ' + message };
    }
  }

  return { success: true, message: '不正日付を正しく拒否' };
}

function testValidateAggregateDateRangeRejectsReverseRange() {
  if (typeof parseAndValidateAggregateDateRange !== 'function') {
    return { success: false, message: 'parseAndValidateAggregateDateRange関数が見つかりません' };
  }

  try {
    parseAndValidateAggregateDateRange('2026-04-01', '2026-03-31');
    return { success: false, message: '日付逆転を検出できませんでした' };
  } catch (error) {
    const message = error && error.message ? error.message : String(error || '');
    if (message.indexOf('開始日は終了日以前の日付を指定してください。') === -1) {
      return { success: false, message: '期待外のエラーメッセージ: ' + message };
    }
  }

  return { success: true, message: '日付逆転を正しく拒否' };
}

function testValidateAggregateDateRangeAcceptsValidRange() {
  if (typeof parseAndValidateAggregateDateRange !== 'function') {
    return { success: false, message: 'parseAndValidateAggregateDateRange関数が見つかりません' };
  }

  try {
    const range = parseAndValidateAggregateDateRange('2025-04-01', '2026-03-31');
    const startDate = range && range.startDate;
    const endDate = range && range.endDate;

    if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
      return { success: false, message: 'Dateオブジェクトが返却されていません' };
    }
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return { success: false, message: '返却値に無効な日付が含まれます' };
    }
    if (startDate > endDate) {
      return { success: false, message: '開始日と終了日の順序が不正です' };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }

  return { success: true, message: '正常な期間を受理' };
}

function testBuildMonthKeysForAggregateAcrossFiscalYear() {
  if (typeof buildMonthKeysForAggregate !== 'function') {
    return { success: false, message: 'buildMonthKeysForAggregate関数が見つかりません' };
  }

  const keys = buildMonthKeysForAggregate(new Date(2025, 3, 1), new Date(2026, 2, 31));
  if (!Array.isArray(keys) || keys.length !== 12) {
    return { success: false, message: '月キー数が不正です: ' + JSON.stringify(keys) };
  }
  if (keys[0] !== '2025-04' || keys[keys.length - 1] !== '2026-03') {
    return { success: false, message: '月キー範囲が不正です: ' + JSON.stringify(keys) };
  }

  return { success: true, message: '年度跨ぎの月キー生成を確認' };
}

function testBuildMonthKeysForAggregateSingleMonth() {
  if (typeof buildMonthKeysForAggregate !== 'function') {
    return { success: false, message: 'buildMonthKeysForAggregate関数が見つかりません' };
  }

  const keys = buildMonthKeysForAggregate(new Date(2025, 8, 1), new Date(2025, 8, 30));
  if (!Array.isArray(keys) || keys.length !== 1 || keys[0] !== '2025-09') {
    return { success: false, message: '単月キー生成が不正です: ' + JSON.stringify(keys) };
  }

  return { success: true, message: '単月の月キー生成を確認' };
}

function testCaptureExistingModValuesByMonth() {
  if (typeof captureExistingModValuesByMonth !== 'function') {
    return { success: false, message: 'captureExistingModValuesByMonth関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('tmp_mod_capture_' + Date.now());

  try {
    tempSheet.getRange(4, 1, 2, 1).setNumberFormat('@');
    tempSheet.getRange(25, 1, 2, 1).setNumberFormat('@');

    tempSheet.getRange(4, 1, 2, 1).setValues([
      ['2025-04'],
      ['2025-05']
    ]);
    tempSheet.getRange(4, 18, 2, 1).setValues([
      [1.5],
      [2]
    ]);

    tempSheet.getRange(25, 1, 2, 1).setValues([
      ['2025-04'],
      ['2025-05']
    ]);
    tempSheet.getRange(25, 18, 2, 1).setValues([
      [3],
      [3.5]
    ]);

    const map = captureExistingModValuesByMonth(
      tempSheet,
      ['2025-04', '2025-05'],
      [1, 2],
      21,
      18
    );

    if (!map || !map[1] || !map[2]) {
      return { success: false, message: '退避結果構造が不正です: ' + JSON.stringify(map) };
    }
    if (map[1]['2025-04'] !== 1.5 || map[1]['2025-05'] !== 2) {
      return { success: false, message: '1年退避データが不正です: ' + JSON.stringify(map[1]) };
    }
    if (map[2]['2025-04'] !== 3 || map[2]['2025-05'] !== 3.5) {
      return { success: false, message: '2年退避データが不正です: ' + JSON.stringify(map[2]) };
    }

    return { success: true, message: '既存MOD値の退避を確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    ss.deleteSheet(tempSheet);
  }
}

function testGetModuleActualUnitsForMonth() {
  if (typeof getModuleActualUnitsForMonth !== 'function') {
    return { success: false, message: 'getModuleActualUnitsForMonth関数が見つかりません' };
  }

  const map = {
    byMonth: {
      '2025-04': {
        1: { actual_units: '1.5' },
        2: { actual_units: 'x' }
      }
    }
  };

  const value1 = getModuleActualUnitsForMonth(map, '2025-04', 1);
  const value2 = getModuleActualUnitsForMonth(map, '2025-04', 2);
  const value3 = getModuleActualUnitsForMonth(map, '2025-05', 1);

  if (value1 !== 1.5) {
    return { success: false, message: '数値文字列変換が不正です: ' + value1 };
  }
  if (value2 !== 0) {
    return { success: false, message: '非数値フォールバックが不正です: ' + value2 };
  }
  if (value3 !== 0) {
    return { success: false, message: '月未存在時の戻り値が不正です: ' + value3 };
  }

  return { success: true, message: 'MOD実績取得のフォールバックを確認' };
}

// ========================================
// 簡易テスト（メニュー用）
// ========================================

/**
 * 重要機能のみの簡易テスト
 * 実行時間を短縮したい場合はこちらを使用
 */
function runQuickTest() {
  Logger.clear();
  Logger.log('====================================');
  Logger.log('ポータルマスター 簡易テスト');
  Logger.log('====================================\n');

  const results = { total: 0, passed: 0, failed: 0, skipped: 0, errors: [] };
  runTestGroups_(results, getQuickTestPlan_());

  // 結果表示
  Logger.log('\n====================================');
  Logger.log('簡易テスト結果');
  Logger.log('====================================');
  Logger.log('総テスト数: ' + results.total);
  Logger.log('✅ 成功: ' + results.passed);
  Logger.log('❌ 失敗: ' + results.failed);

  const successRate = results.total > 0 ? Math.round((results.passed / results.total) * 100) : 0;
  Logger.log('成功率: ' + successRate + '%');

  if (results.failed === 0) {
    Logger.log('\n🎉 簡易テスト成功！');
    SpreadsheetApp.getUi().alert('✅ 簡易テスト成功', '成功率: ' + successRate + '%\n詳細はログを確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    Logger.log('\n⚠️  一部失敗あり');
    SpreadsheetApp.getUi().alert('⚠️ 簡易テスト失敗あり', '成功率: ' + successRate + '%\n詳細はログを確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
