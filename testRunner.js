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

  // UIにも結果を表示（エディタ直接実行時はUIコンテキストがないためスキップ）
  try {
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
  } catch (e) {
    Logger.log('[INFO] UIコンテキストなし — ダイアログ表示をスキップしました。');
  }

  hideInternalSheetsAfterTest_();
}

function hideInternalSheetsAfterTest_() {
  try {
    hideSheetForNormalUse_(MODULE_SHEET_NAMES.CONTROL);
    hideSheetForNormalUse_(SETTINGS_SHEET_NAME);
  } catch (error) {
    Logger.log('[WARNING] テスト後の内部シート非表示化に失敗: ' + error.toString());
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
        { name: '2-1. モジュールシート初期化確認', fn: testInitializeModuleSheets },
        { name: '2-2. 累計時数へのMOD統合確認', fn: testModuleCumulativeIntegration },
        { name: '2-3. 表示フォーマット関数確認', fn: testModuleDisplayFormatter },
        { name: '2-4. 45分換算関数確認', fn: testSessionsToUnits },
        { name: '2-5. 旧スキーマ行の置換除外', fn: testReplaceRowsDropsLegacyFiscalRows },
        { name: '2-6. 他年度旧スキーマ行の保持', fn: testReplaceRowsKeepsOtherLegacyFiscalRows },
        { name: '2-7. Bresenham配分均等性', fn: testDistributeByBresenham },
        { name: '2-8. セッション曜日別配分', fn: testAllocateSessionsToDateKeys }
      ]
    },
    {
      title: '【フェーズ3】学年別集計・データ処理',
      tests: [
        { name: '3-1. 日付マップ作成', fn: testCreateDateMap },
        { name: '3-2. 重複日付の先頭行マッピング', fn: testCreateDateMapKeepsFirstRow },
        { name: '3-3. 集計期間バリデーション（不正日付）', fn: testValidateAggregateDateRangeRejectsInvalidDate },
        { name: '3-4. 集計期間バリデーション（日付順）', fn: testValidateAggregateDateRangeRejectsReverseRange },
        { name: '3-5. 集計期間バリデーション（正常系）', fn: testValidateAggregateDateRangeAcceptsValidRange },
        { name: '3-6. 月キー生成（年度跨ぎ）', fn: testBuildMonthKeysForAggregateAcrossFiscalYear },
        { name: '3-7. 月キー生成（単月）', fn: testBuildMonthKeysForAggregateSingleMonth },
        { name: '3-8. 既存MOD値の月別退避', fn: testCaptureExistingModValuesByMonth },
        { name: '3-9. MOD実績取得関数', fn: testGetModuleActualUnitsForMonth }
      ]
    },
    {
      title: '【フェーズ4】設定・バリデーション',
      tests: [
        { name: '4-1. トリガー設定値読み込み', fn: testGetTriggerSettings },
        { name: '4-2. トリガー設定バリデーション', fn: testValidateTriggerSettings },
        { name: '4-3. トリガー設定正規化', fn: testNormalizeTriggerSettings },
        { name: '4-4. 年度更新設定バリデーション', fn: testValidateAnnualUpdateSettings }
      ]
    },
    {
      title: '【フェーズ5】共通関数',
      tests: [
        { name: '5-1. 日付フォーマット関数', fn: testFormatDateToJapanese },
        { name: '5-2. 名前抽出関数', fn: testExtractFirstName },
        { name: '5-3. 日付正規化関数', fn: testNormalizeToDate },
        { name: '5-4. カレンダー日付範囲抽出', fn: testExtractDateRangeFromData }
      ]
    },
    {
      title: '【フェーズ6】運用導線（非破壊）',
      tests: [
        { name: '6-1. 設定シート非表示動作', fn: testSettingsSheetHiddenForNormalUse },
        { name: '6-2. 年度更新設定ダイアログ定義', fn: testAnnualUpdateDialogDefinition },
        { name: '6-3. 自動トリガー設定ダイアログ定義', fn: testTriggerSettingsDialogDefinition },
        { name: '6-4. 年度更新安全性パターン', fn: testCopyAndClearSafetyPattern },
        { name: '6-5. カレンダー同期管理マーカー', fn: testSyncCalendarsManagedMarkerPattern }
      ]
    },
    {
      title: '【フェーズ7】コード品質・ロジック検証',
      tests: [
        { name: '7-1. var宣言ゼロ検証', fn: testNoVarDeclarations },
        { name: '7-2. ログプレフィックス標準化', fn: testLogPrefixStandard },
        { name: '7-3. エラーハンドリング完備', fn: testErrorHandlingPresence },
        { name: '7-4. XSS安全性確認', fn: testOpenWeeklyReportFolderXssSafe },
        { name: '7-5. 累計カテゴリ導出確認', fn: testCumulativeCategoriesDerivedFromEventCategories },
        { name: '7-6. 日付変換ヘルパー', fn: testConvertCellValue },
        { name: '7-7. 日付行検索', fn: testFindDateRow },
        { name: '7-8. イベント時間解析', fn: testParseEventTimesAndDates },
        { name: '7-9. 累計計算ロジック', fn: testCalculateResultsForGrade },
        { name: '7-10. 月キー正規化', fn: testNormalizeAggregateMonthKey },
        { name: '7-11. 名前結合関数', fn: testJoinNamesWithNewline },
        { name: '7-12. 全角半角変換', fn: testConvertFullWidthToHalfWidth },
        { name: '7-13. 分解析関数', fn: testParseMinute },
        { name: '7-14. バッチ読み取り確認', fn: testAssignDutyBatchReads },
        { name: '7-15. カレンダーイベントキー生成', fn: testBuildCalendarEventKey }
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
        { name: 'Q-8. Bresenham配分均等性', fn: testDistributeByBresenham }
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

/**
 * 累計時数統合の検証（非破壊テスト）
 * 本番シートへの書き込みを行わず、以下を検証する:
 * 1. 統合関数の存在確認
 * 2. 純粋な計算ロジック（buildGradeTotalsFromDailyAndExceptions）の検証
 * 3. 累計時数シートの構造を読み取り専用で確認
 */
function testModuleCumulativeIntegration() {
  if (typeof syncModuleHoursWithCumulative !== 'function') {
    return { success: false, message: 'syncModuleHoursWithCumulative関数が見つかりません' };
  }
  if (typeof buildGradeTotalsFromDailyAndExceptions !== 'function') {
    return { success: false, message: 'buildGradeTotalsFromDailyAndExceptions関数が見つかりません' };
  }

  // 計算ロジックをモックデータで検証（副作用なし）
  const mockDailyTotals = {};
  const mockExceptionTotals = { byGrade: {}, thisWeekByGrade: {} };
  for (let grade = MODULE_GRADE_MIN; grade <= MODULE_GRADE_MAX; grade++) {
    mockDailyTotals[grade] = { plannedSessions: 21, elapsedSessions: 15, thisWeekSessions: 3 };
    mockExceptionTotals.byGrade[grade] = 3;
    mockExceptionTotals.thisWeekByGrade[grade] = 1;
  }

  const gradeTotals = buildGradeTotalsFromDailyAndExceptions(mockDailyTotals, mockExceptionTotals);
  const grade1 = gradeTotals[1];
  if (grade1.actualSessions !== 18) {
    return { success: false, message: '実施セッション計算が不正: 期待18, 実際' + grade1.actualSessions };
  }
  if (grade1.thisWeekSessions !== 4) {
    return { success: false, message: '今週セッション計算が不正: 期待4, 実際' + grade1.thisWeekSessions };
  }

  // 累計時数シートの構造を読み取り専用で確認
  const cumulativeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUMULATIVE_SHEET.NAME);
  if (!cumulativeSheet) {
    return { success: false, message: '累計時数シートが見つかりません' };
  }

  const lastCol = cumulativeSheet.getLastColumn();
  if (lastCol < MODULE_CUMULATIVE_COLUMNS.PLAN) {
    return { success: false, message: 'MOD列が存在しません（最終列: ' + lastCol + '）' };
  }

  const headers = cumulativeSheet.getRange(2, MODULE_CUMULATIVE_COLUMNS.PLAN, 1, 3).getValues()[0];
  const expectedHeaders = ['MOD計画累計', 'MOD実施累計', 'MOD差分'];
  const mismatch = expectedHeaders.filter(function(header, index) {
    return headers[index] !== header;
  });

  if (mismatch.length > 0) {
    return { success: false, message: 'MOD列ヘッダーが不正: ' + JSON.stringify(headers) };
  }

  const displayHeaderRow = cumulativeSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  if (displayHeaderRow.indexOf(MODULE_DISPLAY_HEADER) === -1) {
    return { success: false, message: MODULE_DISPLAY_HEADER + '列が見つかりません' };
  }

  return { success: true, message: '計算ロジック検証 + 累計時数シート構造確認（読み取り専用）' };
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
// 設定・バリデーション
// ========================================

function testGetTriggerSettings() {
  try {
    const settings = getTriggerSettings();

    if (!settings || typeof settings !== 'object') {
      return { success: false, message: '設定オブジェクトが取得できません' };
    }

    const requiredSections = ['weeklyPdf', 'cumulativeHours', 'calendarSync', 'dailyLink'];
    const missingSections = requiredSections.filter(function(section) {
      return !Object.prototype.hasOwnProperty.call(settings, section);
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

function testNormalizeToDate() {
  if (typeof normalizeToDate !== 'function') {
    return { success: false, message: 'normalizeToDate関数が見つかりません' };
  }

  // Date入力: 時刻が00:00:00にリセットされること
  const fromDate = normalizeToDate(new Date(2025, 3, 1, 14, 30, 45));
  if (!fromDate || fromDate.getFullYear() !== 2025 || fromDate.getMonth() !== 3 || fromDate.getDate() !== 1) {
    return { success: false, message: 'Date入力の日付部分が不正' };
  }
  if (fromDate.getHours() !== 0 || fromDate.getMinutes() !== 0 || fromDate.getSeconds() !== 0) {
    return { success: false, message: 'Date入力の時刻リセットが不正' };
  }

  // yyyy-MM-dd文字列
  const fromString = normalizeToDate('2025-04-01');
  if (!fromString || fromString.getFullYear() !== 2025 || fromString.getMonth() !== 3 || fromString.getDate() !== 1) {
    return { success: false, message: 'yyyy-MM-dd文字列パースが不正: ' + fromString };
  }

  // null/undefined/空文字 → null
  if (normalizeToDate(null) !== null || normalizeToDate(undefined) !== null || normalizeToDate('') !== null) {
    return { success: false, message: '空値がnullを返しません' };
  }

  // 不正文字列 → null
  if (normalizeToDate('invalid-date-string') !== null) {
    return { success: false, message: '不正文字列がnullを返しません' };
  }

  return { success: true, message: '4パターンの日付正規化を確認' };
}

function testExtractDateRangeFromData() {
  if (typeof extractDateRangeFromData_ !== 'function') {
    return { success: false, message: 'extractDateRangeFromData_関数が見つかりません' };
  }

  // 正常系: ヘッダー行 + データ行。DATE_INDEXは1（B列相当）
  const data = [
    ['header', 'date_header'],
    ['row1', new Date(2025, 5, 15)],
    ['row2', new Date(2025, 3, 1)],
    ['row3', new Date(2025, 11, 31)]
  ];

  const range = extractDateRangeFromData_(data);
  if (!range || !range.minDate || !range.maxDate) {
    return { success: false, message: '日付範囲が取得できません' };
  }
  if (range.minDate.getMonth() !== 3 || range.maxDate.getMonth() !== 11) {
    return { success: false, message: '最小/最大日付が不正: min=' + range.minDate + ', max=' + range.maxDate };
  }

  // 空データ → null
  const emptyRange = extractDateRangeFromData_([['header', 'date_header']]);
  if (emptyRange !== null) {
    return { success: false, message: '空データでnullが返りません' };
  }

  return { success: true, message: '正常系・空データの日付範囲抽出を確認' };
}

// ========================================
// データ処理テスト
// ========================================

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
    const html = HtmlService.createTemplateFromFile('annualUpdateSettingsDialog').evaluate();
    const content = html.getContent();
    if (!content || content.length === 0) {
      return { success: false, message: '年度更新設定ダイアログHTMLが空です' };
    }
    return { success: true, message: '年度更新設定ダイアログHTML（テンプレート評価）を確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testTriggerSettingsDialogDefinition() {
  if (typeof showTriggerSettingsDialog !== 'function') {
    return { success: false, message: 'showTriggerSettingsDialog関数が見つかりません' };
  }

  try {
    const html = HtmlService.createTemplateFromFile('triggerSettingsDialog').evaluate();
    const content = html.getContent();
    if (!content || content.length === 0) {
      return { success: false, message: '自動トリガー設定ダイアログHTMLが空です' };
    }
    return { success: true, message: '自動トリガー設定ダイアログHTML（テンプレート評価）を確認' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ========================================
// フェーズ7: コード品質・ロジック検証テスト
// ========================================

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
  if (typeof convertCellValue_ !== 'function') {
    return { success: false, message: 'convertCellValue_関数が見つかりません' };
  }

  const case1 = convertCellValue_(new Date(2025, 3, 1), 2025);
  if (case1 !== '2025/04/01') {
    return { success: false, message: 'Date変換が不正: ' + case1 };
  }

  const case2 = convertCellValue_('4月1日', 2025);
  if (case2 !== '2025/04/01') {
    return { success: false, message: '文字列変換が不正: ' + case2 };
  }

  const case3 = convertCellValue_('', 2025);
  if (case3 !== '') {
    return { success: false, message: '空文字列の処理が不正: ' + case3 };
  }

  const case4 = convertCellValue_(null, 2025);
  if (case4 !== '') {
    return { success: false, message: 'null処理が不正: ' + case4 };
  }

  return { success: true, message: '4ケースの日付変換を確認' };
}

function testFindDateRow() {
  if (typeof findDateRow_ !== 'function') {
    return { success: false, message: 'findDateRow_関数が見つかりません' };
  }

  const testValues = [[''], [new Date(2025, 3, 1)], [new Date(2025, 3, 2)]];
  const result = findDateRow_(testValues, '2025/04/02', 2025);
  if (result !== 3) {
    return { success: false, message: '行検索結果が不正: 期待3, 実際' + result };
  }

  const notFound = findDateRow_(testValues, '2025/05/01', 2025);
  if (notFound !== null) {
    return { success: false, message: '未存在検索がnullを返しません: ' + notFound };
  }

  return { success: true, message: '日付行検索を確認' };
}

function testParseEventTimesAndDates() {
  if (typeof parseEventTimesAndDates_ !== 'function') {
    return { success: false, message: 'parseEventTimesAndDates_関数が見つかりません' };
  }

  const testDate = new Date(2025, 3, 1);

  const allDay = parseEventTimesAndDates_('入学式', testDate);
  if (!allDay.isAllDay) {
    return { success: false, message: '全日イベント判定が不正' };
  }

  const rangeTime = parseEventTimesAndDates_('会議 10:00~12:00', testDate);
  if (rangeTime.isAllDay) {
    return { success: false, message: '時間範囲イベントが全日扱いされています' };
  }

  const singleTime = parseEventTimesAndDates_('集会 9:00', testDate);
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
  if (typeof parseMinute_ !== 'function') {
    return { success: false, message: 'parseMinute_関数が見つかりません' };
  }

  const cases = [
    { input: '', expected: 0 },
    { input: '半', expected: 30 },
    { input: '30分', expected: 30 },
    { input: '15', expected: 15 },
    { input: null, expected: 0 }
  ];

  for (let i = 0; i < cases.length; i++) {
    const result = parseMinute_(cases[i].input);
    if (result !== cases[i].expected) {
      return { success: false, message: '入力"' + cases[i].input + '": 期待' + cases[i].expected + ', 実際' + result };
    }
  }

  return { success: true, message: cases.length + 'ケースの分解析を確認' };
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

function testDistributeByBresenham() {
  // Case 1: 均等配分（3日に3セッション → 全日に1セッションずつ）
  const dates1 = [
    new Date(2025, 5, 2),
    new Date(2025, 5, 4),
    new Date(2025, 5, 6)
  ];
  const alloc1 = {};
  distributeByBresenham(dates1, 3, alloc1);
  const keys1 = Object.keys(alloc1);
  if (keys1.length !== 3) {
    return { success: false, message: '3日に3セッション: 全日に配分されるべき（実際: ' + keys1.length + '日）' };
  }

  // Case 2: 不均等配分（5日に2セッション → 2日のみ選択）
  const dates2 = [
    new Date(2025, 5, 2),
    new Date(2025, 5, 4),
    new Date(2025, 5, 6),
    new Date(2025, 5, 9),
    new Date(2025, 5, 11)
  ];
  const alloc2 = {};
  distributeByBresenham(dates2, 2, alloc2);
  const keys2 = Object.keys(alloc2);
  if (keys2.length !== 2) {
    return { success: false, message: '5日に2セッション: 2日に配分されるべき（実際: ' + keys2.length + '日）' };
  }

  // Case 3: 0セッション → 配分なし
  const alloc3 = {};
  distributeByBresenham(dates1, 0, alloc3);
  if (Object.keys(alloc3).length !== 0) {
    return { success: false, message: '0セッション: 配分なしであるべき' };
  }

  return { success: true, message: '3ケースのBresenham配分を確認' };
}

function testAllocateSessionsToDateKeys() {
  // デフォルト曜日優先度（月水金）に対応する日付を使用
  const dates = [
    new Date(2025, 5, 2),
    new Date(2025, 5, 4),
    new Date(2025, 5, 6),
    new Date(2025, 5, 9),
    new Date(2025, 5, 11),
    new Date(2025, 5, 13)
  ];

  // Case 1: 正常配分 — 合計がセッション数と一致
  const result = allocateSessionsToDateKeys(3, dates);
  let totalAllocated = 0;
  Object.keys(result.allocations).forEach(function(k) {
    totalAllocated += result.allocations[k];
  });
  if (totalAllocated + result.overflow !== 3) {
    return { success: false, message: '配分合計+溢れ(' + totalAllocated + '+' + result.overflow + ')が入力(3)と不一致' };
  }

  // Case 2: 0セッション → 配分なし・溢れなし
  const resultZero = allocateSessionsToDateKeys(0, dates);
  if (Object.keys(resultZero.allocations).length !== 0 || resultZero.overflow !== 0) {
    return { success: false, message: '0セッション: 配分なし・溢れなしであるべき' };
  }

  // Case 3: 空日付配列 → 全セッションが溢れ
  const resultEmpty = allocateSessionsToDateKeys(5, []);
  if (resultEmpty.overflow !== 5) {
    return { success: false, message: '空日付: 全セッションが溢れるべき（実際: ' + resultEmpty.overflow + '）' };
  }

  return { success: true, message: '3ケースのセッション配分を確認' };
}

function testCopyAndClearSafetyPattern() {
  const source = String(copyAndClear);

  // OK_CANCEL 確認ダイアログの使用
  if (source.indexOf('OK_CANCEL') === -1) {
    return { success: false, message: 'OK_CANCEL確認ダイアログが見つかりません' };
  }

  // バックアップ整合性検証（コピー後にシート存在・行数を確認してからクリア）
  if (source.indexOf('verifiedSheet') === -1 || source.indexOf('getLastRow') === -1) {
    return { success: false, message: 'バックアップ整合性検証が見つかりません' };
  }

  // clearContent使用（deleteRowsではなくデータのみクリア）
  if (source.indexOf('clearContent') === -1) {
    return { success: false, message: 'clearContent使用が確認できません' };
  }

  // LockService による同時実行保護
  if (source.indexOf('LockService') === -1) {
    return { success: false, message: 'LockServiceによる同時実行保護が見つかりません' };
  }

  return { success: true, message: '年度更新安全性パターンを確認（確認ダイアログ・検証・クリア方式・排他制御）' };
}

function testSyncCalendarsManagedMarkerPattern() {
  const source = String(processEventUpdates_);

  // 管理マーカー判定による選択的削除
  if (source.indexOf('isManagedCalendarEvent_') === -1) {
    return { success: false, message: '管理マーカー判定が見つかりません' };
  }

  // 管理イベントのみ削除（ユーザー手動イベントを保護）
  if (source.indexOf('managedExistingEventMap') === -1) {
    return { success: false, message: '管理イベント限定削除パターンが見つかりません' };
  }

  // 新規イベントへの管理マーカー付与
  if (source.indexOf('markCalendarEventAsManaged_') === -1) {
    return { success: false, message: '新規イベントへの管理マーカー付与が見つかりません' };
  }

  // syncCalendars本体のLockService保護
  const syncSource = String(syncCalendars);
  if (syncSource.indexOf('LockService') === -1) {
    return { success: false, message: 'syncCalendarsにLockServiceによる排他制御が見つかりません' };
  }

  return { success: true, message: 'カレンダー管理マーカーパターンを確認（判定・限定削除・マーカー付与・排他制御）' };
}

function testBuildCalendarEventKey() {
  const start = new Date(2025, 3, 1, 9, 0, 0);
  const end = new Date(2025, 3, 1, 10, 0, 0);

  // Case 1: 同一イベント → 同一キー
  const key1 = buildCalendarEventKey_('入学式', start, end);
  const key2 = buildCalendarEventKey_('入学式', start, end);
  if (key1 !== key2) {
    return { success: false, message: '同一イベントのキーが一致しません' };
  }

  // Case 2: 異なるタイトル → 異なるキー
  const key3 = buildCalendarEventKey_('始業式', start, end);
  if (key1 === key3) {
    return { success: false, message: '異なるタイトルのキーが同一です' };
  }

  // Case 3: 異なる終了時刻 → 異なるキー
  const end2 = new Date(2025, 3, 1, 11, 0, 0);
  const key4 = buildCalendarEventKey_('入学式', start, end2);
  if (key1 === key4) {
    return { success: false, message: '異なる終了時刻のキーが同一です' };
  }

  return { success: true, message: '3ケースのキー一意性を確認' };
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
  } else {
    Logger.log('\n⚠️  一部失敗あり');
  }

  try {
    const ui = SpreadsheetApp.getUi();
    if (results.failed === 0) {
      ui.alert('✅ 簡易テスト成功', '成功率: ' + successRate + '%\n詳細はログを確認してください。', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ 簡易テスト失敗あり', '成功率: ' + successRate + '%\n詳細はログを確認してください。', ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log('[INFO] UIコンテキストなし — ダイアログ表示をスキップしました。');
  }

  hideInternalSheetsAfterTest_();
}
