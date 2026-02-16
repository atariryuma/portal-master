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

  hideInternalSheetsAfterTest_();

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
        { name: '2-6. 表示列の固定列定数確認', fn: testModuleDisplayColumnIsFixed },
        { name: '2-7. 実施曜日フィルタデフォルト', fn: testWeekdayFilterDefault },
        { name: '2-8. 実施曜日パース', fn: testWeekdayFilterParsing },
        { name: '2-9. 曜日シリアライズ', fn: testSerializeWeekdays }
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
        { name: '3-8. 月キー生成（年度跨ぎ）', fn: testListMonthKeysInRangeAcrossFiscalYear },
        { name: '3-9. 月キー生成（単月）', fn: testListMonthKeysInRangeSingleMonth },
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
        { name: '5-2. 名前抽出関数', fn: testExtractFirstName }
      ]
    },
    {
      title: '【フェーズ6】運用導線（非破壊）',
      tests: [
        { name: '6-1. 設定シート非表示動作', fn: testSettingsSheetHiddenForNormalUse },
        { name: '6-2. 年度更新設定ダイアログ定義', fn: testAnnualUpdateDialogDefinition },
        { name: '6-3. 自動トリガー設定ダイアログ定義', fn: testTriggerSettingsDialogDefinition },
        { name: '6-4. 年間行事インポート導線定義', fn: testImportAnnualEventsDefinition },
        { name: '6-5. onOpen設定シート非表示配線', fn: testOnOpenWiresSettingsSheetHide },
        { name: '6-6. 年度更新現行ファイルクリア配線', fn: testCopyAndClearTargetsActiveFileAfterCopy }
      ]
    },
    {
      title: '【フェーズ7】最適化検証',
      tests: [
        { name: '7-1. マジックナンバー定数確認', fn: testMagicNumberConstants },
        { name: '7-2. var宣言ゼロ検証', fn: testNoVarDeclarations },
        { name: '7-3. ログプレフィックス標準化', fn: testLogPrefixStandard },
        { name: '7-4. エラーハンドリング完備', fn: testErrorHandlingPresence },
        { name: '7-5. XSS安全性確認', fn: testOpenWeeklyReportFolderXssSafe },
        { name: '7-6. 累計カテゴリ導出確認', fn: testCumulativeCategoriesDerivedFromEventCategories },
        { name: '7-7. 日付変換ヘルパー', fn: testConvertCellValue },
        { name: '7-8. 日付行検索', fn: testFindDateRow },
        { name: '7-9. イベント時間解析', fn: testParseEventTimesAndDates },
        { name: '7-10. 累計計算ロジック', fn: testCalculateResultsForGrade },
        { name: '7-11. 月キー正規化', fn: testNormalizeAggregateMonthKey },
        { name: '7-12. 名前結合関数', fn: testJoinNamesWithNewline },
        { name: '7-13. 全角半角変換', fn: testConvertFullWidthToHalfWidth },
        { name: '7-14. 分解析関数', fn: testParseMinute },
        { name: '7-15. 公開関数定義確認', fn: testPublicFunctionDefinitions },
        { name: '7-16. バッチ読み取り確認', fn: testAssignDutyBatchReads },
        { name: '7-17. 重複コード排除確認', fn: testNoDuplicateDateFormatter },
        { name: '7-18. moduleHours分割確認', fn: testModuleHoursDecomposition }
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
        { name: 'Q-6. 既存MOD値の月別退避', fn: testCaptureExistingModValuesByMonth },
        { name: 'Q-7. 設定シート非表示動作', fn: testSettingsSheetHiddenForNormalUse },
        { name: 'Q-8. 年度更新現行ファイルクリア配線', fn: testCopyAndClearTargetsActiveFileAfterCopy }
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
  const requiredSheets = ['マスター', '時数様式'];
  const missingSheets = [];

  requiredSheets.forEach(function(sheetName) {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  });

  if (missingSheets.length > 0) {
    return { success: false, message: '不足シート: ' + missingSheets.join(', ') };
  }

  try {
    getSettingsSheetOrThrow();
  } catch (error) {
    return { success: false, message: '設定シート（' + SETTINGS_SHEET_NAME + '）が見つかりません' };
  }

  return { success: true, message: (requiredSheets.length + 1) + '個の必須シートを確認' };
}

function testConfigSheetStructure() {
  let sheet;
  try {
    sheet = getSettingsSheetOrThrow();
  } catch (error) {
    return { success: false, message: '設定シート（' + SETTINGS_SHEET_NAME + '）が見つかりません' };
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
  const requiredConstantsMap = {
    'MODULE_SHEET_NAMES': typeof MODULE_SHEET_NAMES !== 'undefined' ? MODULE_SHEET_NAMES : undefined,
    'MODULE_SETTING_KEYS': typeof MODULE_SETTING_KEYS !== 'undefined' ? MODULE_SETTING_KEYS : undefined,
    'MODULE_DATA_VERSION': typeof MODULE_DATA_VERSION !== 'undefined' ? MODULE_DATA_VERSION : undefined,
    'MODULE_FISCAL_YEAR_START_MONTH': typeof MODULE_FISCAL_YEAR_START_MONTH !== 'undefined' ? MODULE_FISCAL_YEAR_START_MONTH : undefined,
    'MODULE_CUMULATIVE_COLUMNS': typeof MODULE_CUMULATIVE_COLUMNS !== 'undefined' ? MODULE_CUMULATIVE_COLUMNS : undefined
  };

  const missingConstants = Object.keys(requiredConstantsMap).filter(function(constantName) {
    return typeof requiredConstantsMap[constantName] === 'undefined';
  });

  if (missingConstants.length > 0) {
    return { success: false, message: '不足定数: ' + missingConstants.join(', ') };
  }

  if (MODULE_FISCAL_YEAR_START_MONTH !== 4) {
    return { success: false, message: '年度開始月が4月固定になっていません' };
  }

  return { success: true, message: Object.keys(requiredConstantsMap).length + '個のモジュール定数を確認' };
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
    // calculateCumulativeHoursと同じ基準日を使用して、テスト実行による副作用を最小化
    syncModuleHoursWithCumulative(getCurrentOrNextSaturday());
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

function testModuleDisplayColumnIsFixed() {
  if (typeof MODULE_CUMULATIVE_COLUMNS === 'undefined' ||
      typeof MODULE_CUMULATIVE_COLUMNS.DISPLAY === 'undefined') {
    return { success: false, message: 'MODULE_CUMULATIVE_COLUMNS.DISPLAY が定義されていません' };
  }

  if (MODULE_CUMULATIVE_COLUMNS.DISPLAY !== 16) {
    return { success: false, message: '表示列が16(P列)ではありません: ' + MODULE_CUMULATIVE_COLUMNS.DISPLAY };
  }

  if (typeof breakMergesInRange !== 'function') {
    return { success: false, message: 'breakMergesInRange関数が見つかりません' };
  }

  if (typeof cleanupStaleDisplayColumns !== 'function') {
    return { success: false, message: 'cleanupStaleDisplayColumns関数が見つかりません' };
  }

  return { success: true, message: 'MOD表示列の固定列定数と補助関数を確認' };
}

function testWeekdayFilterDefault() {
  if (typeof getEnabledWeekdays !== 'function') {
    return { success: false, message: 'getEnabledWeekdays関数が見つかりません' };
  }

  const result = getEnabledWeekdays({});
  if (!Array.isArray(result) || result.length !== 3) {
    return { success: false, message: 'デフォルト曜日が[1,3,5]ではありません: ' + JSON.stringify(result) };
  }
  if (result[0] !== 1 || result[1] !== 3 || result[2] !== 5) {
    return { success: false, message: 'デフォルト曜日の値が不正: ' + JSON.stringify(result) };
  }

  return { success: true, message: 'デフォルト実施曜日（月水金）を確認' };
}

function testWeekdayFilterParsing() {
  if (typeof getEnabledWeekdays !== 'function') {
    return { success: false, message: 'getEnabledWeekdays関数が見つかりません' };
  }

  const result1 = getEnabledWeekdays({ WEEKDAYS_ENABLED: '1,2,4' });
  if (result1.length !== 3 || result1[0] !== 1 || result1[1] !== 2 || result1[2] !== 4) {
    return { success: false, message: '曜日パース結果が不正: ' + JSON.stringify(result1) };
  }

  const result2 = getEnabledWeekdays({ WEEKDAYS_ENABLED: 'invalid' });
  if (result2.length !== 3 || result2[0] !== 1 || result2[1] !== 3 || result2[2] !== 5) {
    return { success: false, message: '不正値のフォールバックが不正: ' + JSON.stringify(result2) };
  }

  return { success: true, message: '実施曜日パースロジックを確認' };
}

function testSerializeWeekdays() {
  if (typeof serializeWeekdays !== 'function') {
    return { success: false, message: 'serializeWeekdays関数が見つかりません' };
  }

  const result = serializeWeekdays([5, 1, 3]);
  if (result !== '1,3,5') {
    return { success: false, message: 'シリアライズ結果が不正: ' + result };
  }

  const result2 = serializeWeekdays([]);
  if (result2 !== '1,3,5') {
    return { success: false, message: '空配列のフォールバックが不正: ' + result2 };
  }

  return { success: true, message: '曜日シリアライズを確認' };
}

// ========================================
// 設定・バリデーション
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
// 共通関数テスト
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

  for (let i = 0; i < testCases.length; i++) {
    const result = extractFirstName(testCases[i].input);
    if (result !== testCases[i].expected) {
      return { success: false, message: '入力: ' + testCases[i].input + ', 期待値: ' + testCases[i].expected + ', 実際: ' + result };
    }
  }

  return { success: true, message: testCases.length + '件のテストケースが成功' };
}

// ========================================
// データ処理テスト
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

function testListMonthKeysInRangeAcrossFiscalYear() {
  if (typeof listMonthKeysInRange !== 'function') {
    return { success: false, message: 'listMonthKeysInRange関数が見つかりません' };
  }

  const keys = listMonthKeysInRange(new Date(2025, 3, 1), new Date(2026, 2, 31));
  if (!Array.isArray(keys) || keys.length !== 12) {
    return { success: false, message: '月キー数が不正です: ' + JSON.stringify(keys) };
  }
  if (keys[0] !== '2025-04' || keys[keys.length - 1] !== '2026-03') {
    return { success: false, message: '月キー範囲が不正です: ' + JSON.stringify(keys) };
  }

  return { success: true, message: '年度跨ぎの月キー生成を確認' };
}

function testListMonthKeysInRangeSingleMonth() {
  if (typeof listMonthKeysInRange !== 'function') {
    return { success: false, message: 'listMonthKeysInRange関数が見つかりません' };
  }

  const keys = listMonthKeysInRange(new Date(2025, 8, 1), new Date(2025, 8, 30));
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

function testSettingsSheetHiddenForNormalUse() {
  if (typeof hideSheetForNormalUse_ !== 'function') {
    return { success: false, message: 'hideSheetForNormalUse_関数が見つかりません' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    return { success: false, message: SETTINGS_SHEET_NAME + 'シートが見つかりません' };
  }

  const wasHidden = settingsSheet.isSheetHidden();
  const activeSheet = ss.getActiveSheet();
  const activeSheetId = activeSheet ? activeSheet.getSheetId() : null;
  const visibleCount = ss.getSheets().filter(function(sheet) {
    return !sheet.isSheetHidden();
  }).length;

  if (!wasHidden && visibleCount <= 1) {
    return { skip: true, message: '表示中シートが1枚のみのため非表示テストをスキップ' };
  }

  try {
    hideSheetForNormalUse_(SETTINGS_SHEET_NAME);
    if (!settingsSheet.isSheetHidden()) {
      return { success: false, message: SETTINGS_SHEET_NAME + 'シートが非表示になりません' };
    }
    return {
      success: true,
      message: wasHidden ? '既に非表示状態を確認' : '非表示化動作を確認（テスト後に元へ復元）'
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (!wasHidden && settingsSheet.isSheetHidden()) {
      settingsSheet.showSheet();

      if (activeSheetId !== null) {
        const originalActiveSheet = ss.getSheets().find(function(sheet) {
          return sheet.getSheetId() === activeSheetId;
        });
        if (originalActiveSheet && !originalActiveSheet.isSheetHidden()) {
          ss.setActiveSheet(originalActiveSheet);
        }
      }
    }
  }
}

function testAnnualUpdateDialogDefinition() {
  if (typeof showAnnualUpdateSettingsDialog !== 'function') {
    return { success: false, message: 'showAnnualUpdateSettingsDialog関数が見つかりません' };
  }

  try {
    const html = HtmlService.createHtmlOutputFromFile('annualUpdateSettingsDialog');
    const content = html.getContent();
    if (!content || content.length === 0) {
      return { success: false, message: '年度更新設定ダイアログHTMLが空です' };
    }
    return { success: true, message: '年度更新設定ダイアログHTMLを確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testTriggerSettingsDialogDefinition() {
  if (typeof showTriggerSettingsDialog !== 'function') {
    return { success: false, message: 'showTriggerSettingsDialog関数が見つかりません' };
  }

  try {
    const html = HtmlService.createHtmlOutputFromFile('triggerSettingsDialog');
    const content = html.getContent();
    if (!content || content.length === 0) {
      return { success: false, message: '自動トリガー設定ダイアログHTMLが空です' };
    }
    return { success: true, message: '自動トリガー設定ダイアログHTMLを確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testImportAnnualEventsDefinition() {
  if (typeof importAnnualEvents !== 'function') {
    return { success: false, message: 'importAnnualEvents関数が見つかりません' };
  }

  const source = String(importAnnualEvents);
  const requiredFragments = [
    'getSettingsSheetOrThrow',
    'ANNUAL_UPDATE_CONFIG_CELLS.BASE_SUNDAY',
    'SpreadsheetApp.openByUrl'
  ];

  const missingFragments = requiredFragments.filter(function(fragment) {
    return source.indexOf(fragment) === -1;
  });

  if (missingFragments.length > 0) {
    return { success: false, message: '導線コード不足: ' + missingFragments.join(', ') };
  }

  if (source.indexOf('年度更新作業') !== -1) {
    return { success: false, message: '旧設定シート名参照が残っています' };
  }

  return { success: true, message: '年間行事インポート導線を確認' };
}

function testOnOpenWiresSettingsSheetHide() {
  if (typeof onOpen !== 'function') {
    return { success: false, message: 'onOpen関数が見つかりません' };
  }

  const source = String(onOpen);
  if (source.indexOf('hideInternalSheetsForNormalUse_') === -1) {
    return { success: false, message: 'onOpenの内部シート非表示配線が見つかりません' };
  }

  const helperSource = String(hideInternalSheetsForNormalUse_);
  if (helperSource.indexOf('MODULE_SHEET_NAMES') === -1 || helperSource.indexOf('SETTINGS_SHEET_NAME') === -1) {
    return { success: false, message: 'hideInternalSheetsForNormalUse_にmodule_control・設定シートが含まれていません' };
  }

  return { success: true, message: 'onOpenの設定シート非表示配線を確認' };
}

function testCopyAndClearTargetsActiveFileAfterCopy() {
  if (typeof copyAndClear !== 'function') {
    return { success: false, message: 'copyAndClear関数が見つかりません' };
  }

  const source = String(copyAndClear);
  const requiredFragments = [
    'makeCopy(',
    "getSheetByName('年間行事予定表')",
    'ANNUAL_SCHEDULE.CLEAR_EVENT_RANGE',
    'ANNUAL_SCHEDULE.CLEAR_DATA_RANGE'
  ];
  const missingFragments = requiredFragments.filter(function(fragment) {
    return source.indexOf(fragment) === -1;
  });

  if (missingFragments.length > 0) {
    return { success: false, message: '現行ファイルクリアの導線不足: ' + missingFragments.join(', ') };
  }

  if (source.indexOf('copiedSheet.getRange(') !== -1 || source.indexOf('SpreadsheetApp.openById') !== -1) {
    return { success: false, message: 'コピー先ファイルを直接クリアする導線が残っています' };
  }

  return { success: true, message: '年度更新はコピー後に現行ファイルをクリアする配線を確認' };
}

// ========================================
// フェーズ7: 最適化検証テスト
// ========================================

function testMagicNumberConstants() {
  const requiredConstantsMap = {
    'MASTER_SHEET': typeof MASTER_SHEET !== 'undefined' ? MASTER_SHEET : undefined,
    'DUTY_ROSTER_SHEET': typeof DUTY_ROSTER_SHEET !== 'undefined' ? DUTY_ROSTER_SHEET : undefined,
    'ANNUAL_SCHEDULE': typeof ANNUAL_SCHEDULE !== 'undefined' ? ANNUAL_SCHEDULE : undefined,
    'JISUU_TEMPLATE': typeof JISUU_TEMPLATE !== 'undefined' ? JISUU_TEMPLATE : undefined,
    'WEEKLY_REPORT': typeof WEEKLY_REPORT !== 'undefined' ? WEEKLY_REPORT : undefined,
    'CUMULATIVE_SHEET': typeof CUMULATIVE_SHEET !== 'undefined' ? CUMULATIVE_SHEET : undefined,
    'IMPORT_CONSTANTS': typeof IMPORT_CONSTANTS !== 'undefined' ? IMPORT_CONSTANTS : undefined
  };

  const missingConstants = Object.keys(requiredConstantsMap).filter(function(name) {
    return typeof requiredConstantsMap[name] === 'undefined';
  });

  if (missingConstants.length > 0) {
    return { success: false, message: '不足定数: ' + missingConstants.join(', ') };
  }

  if (MASTER_SHEET.DUTY_COLUMN !== 41 || ANNUAL_SCHEDULE.DUTY_COLUMN !== 18) {
    return { success: false, message: '定数値が不正です' };
  }

  return { success: true, message: Object.keys(requiredConstantsMap).length + '個の定数グループを確認' };
}

function testNoVarDeclarations() {
  const functionsToCheck = [
    { name: 'importAnnualEvents', fn: importAnnualEvents },
    { name: 'openWeeklyReportFolder', fn: openWeeklyReportFolder },
    { name: 'assignDuty', fn: assignDuty },
    { name: 'updateAnnualDuty', fn: updateAnnualDuty },
    { name: 'updateAnnualEvents', fn: updateAnnualEvents },
    { name: 'countStars', fn: countStars },
    { name: 'saveToPDF', fn: saveToPDF },
    { name: 'setDailyHyperlink', fn: setDailyHyperlink },
    { name: 'breakMergesInRange', fn: breakMergesInRange },
    { name: 'cleanupStaleDisplayColumns', fn: cleanupStaleDisplayColumns }
  ];

  const filesWithVar = [];
  functionsToCheck.forEach(function(item) {
    const source = String(item.fn);
    if (/\bvar\s+/.test(source)) {
      filesWithVar.push(item.name);
    }
  });

  if (filesWithVar.length > 0) {
    return { success: false, message: 'var使用ファイル: ' + filesWithVar.join(', ') };
  }

  return { success: true, message: functionsToCheck.length + '関数でvar不使用を確認' };
}

function testLogPrefixStandard() {
  const functionsToCheck = [
    { name: 'formatDateToJapanese', fn: formatDateToJapanese },
    { name: 'saveToPDF', fn: saveToPDF },
    { name: 'calculateCumulativeHours', fn: calculateCumulativeHours }
  ];

  const unprefixed = [];
  functionsToCheck.forEach(function(item) {
    const source = String(item.fn);
    const logCalls = source.match(/Logger\.log\([^)]+\)/g) || [];
    logCalls.forEach(function(call) {
      if (!/\[(INFO|WARNING|ERROR|DEBUG)\]/.test(call)) {
        unprefixed.push(item.name + ': ' + call.substring(0, 50));
      }
    });
  });

  if (unprefixed.length > 0) {
    return { success: false, message: 'プレフィックスなし: ' + unprefixed.join('; ') };
  }

  return { success: true, message: 'ログプレフィックス標準化を確認' };
}

function testErrorHandlingPresence() {
  const functionsToCheck = [
    { name: 'assignDuty', fn: assignDuty },
    { name: 'updateAnnualDuty', fn: updateAnnualDuty },
    { name: 'countStars', fn: countStars },
    { name: 'setDailyHyperlink', fn: setDailyHyperlink },
    { name: 'saveToPDF', fn: saveToPDF },
    { name: 'openWeeklyReportFolder', fn: openWeeklyReportFolder }
  ];

  const missingTryCatch = [];
  functionsToCheck.forEach(function(item) {
    const source = String(item.fn);
    if (source.indexOf('try') === -1 || source.indexOf('catch') === -1) {
      missingTryCatch.push(item.name);
    }
  });

  if (missingTryCatch.length > 0) {
    return { success: false, message: 'try/catch未実装: ' + missingTryCatch.join(', ') };
  }

  return { success: true, message: functionsToCheck.length + '関数のエラーハンドリングを確認' };
}

function testOpenWeeklyReportFolderXssSafe() {
  const source = String(openWeeklyReportFolder);
  if (source.indexOf('createHtmlOutput') !== -1 && source.indexOf('folderId') !== -1 && source.indexOf('+') !== -1) {
    if (source.indexOf('createTemplate') === -1) {
      return { success: false, message: 'HTML直接連結によるXSSリスクがあります' };
    }
  }

  if (source.indexOf('var ') !== -1) {
    return { success: false, message: 'var宣言が残っています' };
  }

  return { success: true, message: 'XSS安全性とconst/let使用を確認' };
}

function testCumulativeCategoriesDerivedFromEventCategories() {
  if (!Array.isArray(CUMULATIVE_EVENT_CATEGORIES)) {
    return { success: false, message: 'CUMULATIVE_EVENT_CATEGORIESが配列ではありません' };
  }

  const allFromEventCategories = CUMULATIVE_EVENT_CATEGORIES.every(function(cat) {
    return Object.prototype.hasOwnProperty.call(EVENT_CATEGORIES, cat);
  });

  if (!allFromEventCategories) {
    return { success: false, message: 'EVENT_CATEGORIESに含まれないカテゴリがあります' };
  }

  if (CUMULATIVE_EVENT_CATEGORIES.indexOf('補習') !== -1) {
    return { success: false, message: '「補習」が累計対象に含まれています' };
  }

  return { success: true, message: 'EVENT_CATEGORIESからの導出を確認（補習除外）' };
}

function testConvertCellValue() {
  if (typeof convertCellValue !== 'function') {
    return { success: false, message: 'convertCellValue関数が見つかりません' };
  }

  const case1 = convertCellValue(new Date(2025, 3, 1), 2025);
  if (case1 !== '2025/04/01') {
    return { success: false, message: 'Date変換が不正: ' + case1 };
  }

  const case2 = convertCellValue('4月1日', 2025);
  if (case2 !== '2025/04/01') {
    return { success: false, message: '文字列変換が不正: ' + case2 };
  }

  const case3 = convertCellValue('', 2025);
  if (case3 !== '') {
    return { success: false, message: '空文字列の処理が不正: ' + case3 };
  }

  const case4 = convertCellValue(null, 2025);
  if (case4 !== '') {
    return { success: false, message: 'null処理が不正: ' + case4 };
  }

  return { success: true, message: '4ケースの日付変換を確認' };
}

function testFindDateRow() {
  if (typeof findDateRow !== 'function') {
    return { success: false, message: 'findDateRow関数が見つかりません' };
  }

  const testValues = [[''], [new Date(2025, 3, 1)], [new Date(2025, 3, 2)]];
  const result = findDateRow(testValues, '2025/04/02', 2025);
  if (result !== 3) {
    return { success: false, message: '行検索結果が不正: 期待3, 実際' + result };
  }

  const notFound = findDateRow(testValues, '2025/05/01', 2025);
  if (notFound !== null) {
    return { success: false, message: '未存在検索がnullを返しません: ' + notFound };
  }

  return { success: true, message: '日付行検索を確認' };
}

function testParseEventTimesAndDates() {
  if (typeof parseEventTimesAndDates !== 'function') {
    return { success: false, message: 'parseEventTimesAndDates関数が見つかりません' };
  }

  const testDate = new Date(2025, 3, 1);

  const allDay = parseEventTimesAndDates('入学式', testDate);
  if (!allDay.isAllDay) {
    return { success: false, message: '全日イベント判定が不正' };
  }

  const rangeTime = parseEventTimesAndDates('会議 10:00~12:00', testDate);
  if (rangeTime.isAllDay) {
    return { success: false, message: '時間範囲イベントが全日扱いされています' };
  }

  const singleTime = parseEventTimesAndDates('集会 9:00', testDate);
  if (singleTime.isAllDay) {
    return { success: false, message: '単一時間イベントが全日扱いされています' };
  }

  return { success: true, message: '3パターンのイベント時間解析を確認' };
}

function testCalculateResultsForGrade() {
  if (typeof calculateResultsForGrade !== 'function') {
    return { success: false, message: 'calculateResultsForGrade関数が見つかりません' };
  }

  const mockData = [
    ['header', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '○', '○', '', '', '', ''],
    [new Date(2025, 3, 1), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 1, '○', '○', '○', '', '', '']
  ];

  const categories = { '儀式': '儀式' };
  const endDate = new Date(2025, 3, 30);
  const results = calculateResultsForGrade(mockData, 1, endDate, categories);

  if (results["授業時数"] !== 3) {
    return { success: false, message: '授業時数が不正: ' + results["授業時数"] };
  }

  return { success: true, message: '累計計算ロジックを確認' };
}

function testNormalizeAggregateMonthKey() {
  if (typeof normalizeAggregateMonthKey !== 'function') {
    return { success: false, message: 'normalizeAggregateMonthKey関数が見つかりません' };
  }

  const case1 = normalizeAggregateMonthKey(new Date(2025, 3, 15));
  if (case1 !== '2025-04') {
    return { success: false, message: 'Date正規化が不正: ' + case1 };
  }

  const case2 = normalizeAggregateMonthKey('2025-04');
  if (case2 !== '2025-04') {
    return { success: false, message: '文字列正規化が不正: ' + case2 };
  }

  const case3 = normalizeAggregateMonthKey(null);
  if (case3 !== '') {
    return { success: false, message: 'null正規化が不正: ' + case3 };
  }

  return { success: true, message: '3パターンの月キー正規化を確認' };
}

function testJoinNamesWithNewline() {
  if (typeof joinNamesWithNewline !== 'function') {
    return { success: false, message: 'joinNamesWithNewline関数が見つかりません' };
  }

  const case1 = joinNamesWithNewline(['太郎', '花子']);
  if (case1 !== '太郎\n花子') {
    return { success: false, message: '2名結合が不正: ' + JSON.stringify(case1) };
  }

  const case2 = joinNamesWithNewline(['太郎', '', '花子']);
  if (case2 !== '太郎\n花子') {
    return { success: false, message: '空名フィルタが不正: ' + JSON.stringify(case2) };
  }

  const case3 = joinNamesWithNewline([]);
  if (case3 !== '') {
    return { success: false, message: '空配列処理が不正: ' + JSON.stringify(case3) };
  }

  const case4 = joinNamesWithNewline(null);
  if (case4 !== '') {
    return { success: false, message: 'null処理が不正: ' + JSON.stringify(case4) };
  }

  return { success: true, message: '4ケースの名前結合を確認' };
}

function testConvertFullWidthToHalfWidth() {
  if (typeof convertFullWidthToHalfWidth !== 'function') {
    return { success: false, message: 'convertFullWidthToHalfWidth関数が見つかりません' };
  }

  const case1 = convertFullWidthToHalfWidth('１２：３０');
  if (case1 !== '12:30') {
    return { success: false, message: '全角数字変換が不正: ' + case1 };
  }

  const case2 = convertFullWidthToHalfWidth('');
  if (case2 !== '') {
    return { success: false, message: '空文字列処理が不正' };
  }

  const case3 = convertFullWidthToHalfWidth('abc');
  if (case3 !== 'abc') {
    return { success: false, message: '半角文字がそのまま返らない: ' + case3 };
  }

  return { success: true, message: '3ケースの全角半角変換を確認' };
}

function testParseMinute() {
  if (typeof parseMinute !== 'function') {
    return { success: false, message: 'parseMinute関数が見つかりません' };
  }

  const cases = [
    { input: '', expected: 0 },
    { input: '半', expected: 30 },
    { input: '30分', expected: 30 },
    { input: '15', expected: 15 },
    { input: null, expected: 0 }
  ];

  for (let i = 0; i < cases.length; i++) {
    const result = parseMinute(cases[i].input);
    if (result !== cases[i].expected) {
      return { success: false, message: '入力"' + cases[i].input + '": 期待' + cases[i].expected + ', 実際' + result };
    }
  }

  return { success: true, message: cases.length + 'ケースの分解析を確認' };
}

function testPublicFunctionDefinitions() {
  const publicFunctionsMap = {
    'assignDuty': assignDuty,
    'updateAnnualDuty': updateAnnualDuty,
    'updateAnnualEvents': updateAnnualEvents,
    'countStars': countStars,
    'setDailyHyperlink': setDailyHyperlink,
    'saveToPDF': saveToPDF,
    'openWeeklyReportFolder': openWeeklyReportFolder,
    'syncCalendars': syncCalendars,
    'calculateCumulativeHours': calculateCumulativeHours,
    'importAnnualEvents': importAnnualEvents,
    'aggregateSchoolEventsByGrade': aggregateSchoolEventsByGrade,
    'processAggregateSchoolEventsByGrade': processAggregateSchoolEventsByGrade,
    'copyAndClear': copyAndClear,
    'showAnnualUpdateSettingsDialog': showAnnualUpdateSettingsDialog,
    'showTriggerSettingsDialog': showTriggerSettingsDialog,
    'showModulePlanningDialog': showModulePlanningDialog,
    'runAllTests': runAllTests
  };

  const missingFunctions = Object.keys(publicFunctionsMap).filter(function(name) {
    return typeof publicFunctionsMap[name] !== 'function';
  });

  if (missingFunctions.length > 0) {
    return { success: false, message: '未定義関数: ' + missingFunctions.join(', ') };
  }

  return { success: true, message: Object.keys(publicFunctionsMap).length + '個の公開関数を確認' };
}

function testAssignDutyBatchReads() {
  const source = String(assignDuty);

  // ループ内の個別getValue呼び出しがないことを確認
  const hasIndividualReads = /for\s*\([^)]*\)\s*\{[^}]*getRange\([^)]*\)\.getValue\(\)/s.test(source);
  if (hasIndividualReads) {
    return { success: false, message: 'ループ内に個別getValueが残っています' };
  }

  // バッチ読み取りのgetValuesが存在することを確認
  if (source.indexOf('getValues()') === -1) {
    return { success: false, message: 'バッチ読み取り（getValues）が見つかりません' };
  }

  return { success: true, message: 'バッチ読み取りパターンを確認' };
}

function testNoDuplicateDateFormatter() {
  // formatDateRangeForPdf_ は createFileName にインライン化済み
  if (typeof formatDateRangeForPdf_ === 'function') {
    return { success: false, message: '廃止済みのformatDateRangeForPdf_関数が残っています' };
  }

  // createFileName が formatDateToJapanese を再利用していることを確認
  if (typeof createFileName === 'function') {
    const source = String(createFileName);
    if (source.indexOf('formatDateToJapanese') === -1) {
      return { success: false, message: 'createFileNameがformatDateToJapaneseを使用していません' };
    }
  }

  // formatDateRange がまだ存在する場合は重複
  if (typeof formatDateRange === 'function') {
    return { success: false, message: '旧formatDateRange関数が残っています' };
  }

  return { success: true, message: '重複日付フォーマッターなし（createFileNameにインライン化済み）' };
}

function testModuleHoursDecomposition() {
  const expectedFiles = [
    'moduleHoursConstants',
    'moduleHoursDialog',
    'moduleHoursPlanning',
    'moduleHoursControl',
    'moduleHoursDisplay'
  ];

  // 各ファイルからの代表的な関数/定数が存在するか確認
  const checkSymbols = {
    moduleHoursConstants: {
      'MODULE_DEFAULT_ANNUAL_KOMA': typeof MODULE_DEFAULT_ANNUAL_KOMA !== 'undefined',
      'MODULE_CONTROL_MARKERS': typeof MODULE_CONTROL_MARKERS !== 'undefined',
      'MODULE_DEFAULT_WEEKDAYS_ENABLED': typeof MODULE_DEFAULT_WEEKDAYS_ENABLED !== 'undefined',
      'MODULE_WEEKDAY_LABELS': typeof MODULE_WEEKDAY_LABELS !== 'undefined'
    },
    moduleHoursDialog: {
      'showModulePlanningDialog': typeof showModulePlanningDialog === 'function',
      'getModulePlanningDialogState': typeof getModulePlanningDialogState === 'function',
      'saveModuleAnnualTargetFromDialog': typeof saveModuleAnnualTargetFromDialog === 'function',
      'saveModuleSettingsFromDialog': typeof saveModuleSettingsFromDialog === 'function'
    },
    moduleHoursPlanning: {
      'buildDailyPlanFromAnnualTarget': typeof buildDailyPlanFromAnnualTarget === 'function',
      'allocateSessionsToDateKeys': typeof allocateSessionsToDateKeys === 'function'
    },
    moduleHoursControl: {
      'initializeModuleHoursSheetsIfNeeded': typeof initializeModuleHoursSheetsIfNeeded === 'function',
      'readExceptionRows': typeof readExceptionRows === 'function',
      'readModuleSettingsMap': typeof readModuleSettingsMap === 'function'
    },
    moduleHoursDisplay: {
      'syncModuleHoursWithCumulative': typeof syncModuleHoursWithCumulative === 'function',
      'formatSessionsAsMixedFraction': typeof formatSessionsAsMixedFraction === 'function',
      'normalizeToDate': typeof normalizeToDate === 'function',
      'toNumberOrZero': typeof toNumberOrZero === 'function'
    }
  };

  const missing = [];
  Object.keys(checkSymbols).forEach(function(file) {
    Object.keys(checkSymbols[file]).forEach(function(name) {
      if (!checkSymbols[file][name]) {
        missing.push(file + '/' + name);
      }
    });
  });

  if (missing.length > 0) {
    return { success: false, message: '未定義: ' + missing.join(', ') };
  }

  return { success: true, message: expectedFiles.length + 'ファイルの代表関数/定数がすべて定義済み' };
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

  hideInternalSheetsAfterTest_();

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

/**
 * テスト終了後に内部管理シートを非表示に復元
 */
function hideInternalSheetsAfterTest_() {
  try {
    hideInternalSheetsForNormalUse_(true);
  } catch (error) {
    Logger.log('[WARNING] テスト後の内部シート非表示化に失敗: ' + error.toString());
  }
}
