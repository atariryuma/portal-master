/**
 * @fileoverview ウェブアプリ エントリポイント
 * @description doGet/doPostを定義し、ウェブアプリとしてのリクエストを処理します。
 *   今後の拡張（外部連携API、ダッシュボード表示など）の基盤となります。
 */

/**
 * GETリクエストのエントリポイント
 * @param {Object} e - イベントオブジェクト
 * @param {Object} e.parameter - URLクエリパラメータ（単一値）
 * @param {Object} e.parameters - URLクエリパラメータ（配列値）
 * @param {string} e.pathInfo - URLパス情報
 * @return {HtmlOutput|TextOutput} レスポンス
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'home';

  switch (page) {
    case 'status':
      return buildJsonResponse_(getStatusInfo_());
    case 'diagnostics':
      return buildJsonResponse_(runDiagnostics_());
    case 'tests':
      return buildJsonResponse_(runTestsViaWebapp_(e.parameter.suite || 'full', e.parameter.phase || ''));
    default:
      return buildHomePage_();
  }
}

/**
 * POSTリクエストのエントリポイント
 * @param {Object} e - イベントオブジェクト
 * @param {string} e.postData.contents - リクエストボディ
 * @param {string} e.postData.type - Content-Type
 * @return {TextOutput} JSONレスポンス
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || '';

    switch (action) {
      case 'ping':
        return buildJsonResponse_({ success: true, message: 'pong', timestamp: new Date().toISOString() });
      case 'deleteOrphanTriggers':
        return buildJsonResponse_(deleteOrphanTriggers_());
      default:
        return buildJsonResponse_({ success: false, error: '不明なアクション: ' + action }, 400);
    }
  } catch (error) {
    Logger.log('[ERROR] doPost: ' + error.toString());
    return buildJsonResponse_({ success: false, error: 'リクエストの処理に失敗しました' }, 500);
  }
}

// ========================================
// レスポンスビルダー
// ========================================

/**
 * JSONレスポンスを生成する
 * @param {Object} data - レスポンスデータ
 * @return {TextOutput} JSONレスポンス
 */
function buildJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ホームページHTMLを生成する
 * @return {HtmlOutput} HTMLレスポンス
 */
function buildHomePage_() {
  const template = HtmlService.createTemplateFromFile('webappHome');
  return template.evaluate()
    .setTitle('ポータルマスター')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// ステータス情報
// ========================================

/**
 * 主要関数の診断チェックを実行する
 * @return {Object} 診断結果
 */
function runDiagnostics_() {
  const results = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート存在チェック
  const requiredSheets = [
    'マスター', 'app_config', '時数様式', '年間行事予定表',
    '累計時数', '日直表', 'module_control',
    '週報（時数あり）', '週報（時数あり）次週'
  ];
  requiredSheets.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    results.push({
      check: 'シート存在: ' + name,
      status: sheet ? 'OK' : 'MISSING',
      detail: sheet ? '行数: ' + sheet.getLastRow() : 'シートが見つかりません'
    });
  });

  // 主要関数の実行チェック
  const checks = [
    { name: 'getSettingsSheetOrThrow', fn: function() { getSettingsSheetOrThrow(); } },
    { name: 'readModuleSettingsMap', fn: function() { readModuleSettingsMap(); } },
    { name: 'getCurrentOrNextSaturday', fn: function() { getCurrentOrNextSaturday(); } },
    { name: 'getFiscalYear', fn: function() { getFiscalYear(new Date()); } }
  ];

  checks.forEach(function(check) {
    try {
      check.fn();
      results.push({ check: '関数: ' + check.name, status: 'OK', detail: '正常に実行' });
    } catch (error) {
      results.push({ check: '関数: ' + check.name, status: 'ERROR', detail: error.toString() });
    }
  });

  // トリガー一覧
  let triggers = [];
  try {
    triggers = ScriptApp.getProjectTriggers().map(function(t) {
      return {
        function: t.getHandlerFunction(),
        type: t.getEventType().toString(),
        source: t.getTriggerSource().toString()
      };
    });
  } catch (error) {
    triggers = [{ error: error.toString() }];
  }

  const errorCount = results.filter(function(r) { return r.status === 'ERROR' || r.status === 'MISSING'; }).length;

  return {
    success: true,
    timestamp: new Date().toISOString(),
    summary: {
      total: results.length,
      ok: results.length - errorCount,
      errors: errorCount
    },
    results: results,
    triggers: triggers
  };
}

/**
 * テストスイートをウェブアプリ経由で実行する
 * @param {string} suite - 'full' または 'quick'
 * @param {string} phase - フェーズ番号（例: '1', '2'）。空文字で全フェーズ
 * @return {Object} テスト結果
 */
function runTestsViaWebapp_(suite, phase) {
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    skipped: 0,
    errors: [],
    details: []
  };

  let plan = (suite === 'quick') ? getQuickTestPlan_() : getFullTestPlan_();

  // フェーズ指定がある場合はフィルタリング
  if (phase) {
    const phaseNum = parseInt(phase, 10);
    plan = plan.filter(function(group, index) {
      return (index + 1) === phaseNum;
    });
  }

  plan.forEach(function(group) {
    group.tests.forEach(function(testItem) {
      results.total++;
      try {
        const result = testItem.fn();
        if (result.skip) {
          results.skipped++;
          results.details.push({ name: testItem.name, status: 'SKIP', message: result.message });
        } else if (result.success) {
          results.passed++;
          results.details.push({ name: testItem.name, status: 'PASS', message: result.message || '' });
        } else {
          results.failed++;
          results.errors.push(testItem.name + ': ' + result.message);
          results.details.push({ name: testItem.name, status: 'FAIL', message: result.message });
        }
      } catch (error) {
        results.failed++;
        results.errors.push(testItem.name + ': ' + error.toString());
        results.details.push({ name: testItem.name, status: 'ERROR', message: error.toString() });
      }
    });
  });

  const successRate = results.total > 0 ? Math.round((results.passed / results.total) * 100) : 0;

  return {
    success: true,
    suite: suite,
    timestamp: new Date().toISOString(),
    summary: {
      total: results.total,
      passed: results.passed,
      failed: results.failed,
      skipped: results.skipped,
      successRate: successRate + '%'
    },
    errors: results.errors,
    details: results.details
  };
}

/**
 * 対応する関数が存在しない孤立トリガーを削除する
 * @return {Object} 削除結果
 */
function deleteOrphanTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  const deleted = [];

  triggers.forEach(function(trigger) {
    const funcName = trigger.getHandlerFunction();
    try {
      // globalスコープに関数が存在するかチェック
      if (typeof this[funcName] !== 'function') {
        ScriptApp.deleteTrigger(trigger);
        deleted.push(funcName);
        Logger.log('[INFO] 孤立トリガーを削除: ' + funcName);
      }
    } catch (error) {
      Logger.log('[ERROR] トリガー削除失敗: ' + funcName + ' - ' + error.toString());
    }
  });

  return {
    success: true,
    deletedCount: deleted.length,
    deleted: deleted,
    timestamp: new Date().toISOString()
  };
}

/**
 * システムステータス情報を取得する（詳細版）
 * @return {Object} ステータス情報
 */
function getStatusInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート詳細情報
  const sheetDetails = ss.getSheets().map(function(sheet) {
    return {
      name: sheet.getName(),
      rows: sheet.getLastRow(),
      cols: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    };
  });

  // スプレッドシート情報
  const spreadsheetInfo = {
    name: ss.getName(),
    id: ss.getId(),
    url: ss.getUrl(),
    locale: ss.getSpreadsheetLocale(),
    timezone: ss.getSpreadsheetTimeZone(),
    sheetCount: sheetDetails.length
  };

  // GASプロジェクト情報
  const scriptInfo = {
    scriptId: ScriptApp.getScriptId(),
    authMode: ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus().toString()
  };

  // トリガー情報
  const triggers = ScriptApp.getProjectTriggers().map(function(t) {
    return {
      function: t.getHandlerFunction(),
      type: t.getEventType().toString(),
      source: t.getTriggerSource().toString(),
      id: t.getUniqueId()
    };
  });

  // モジュール設定情報
  let moduleSettings = {};
  try {
    moduleSettings = readModuleSettingsMap();
  } catch (e) {
    moduleSettings = { error: e.toString() };
  }

  // app_config情報
  let appConfig = {};
  try {
    const configSheet = ss.getSheetByName('app_config');
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      configData.forEach(function(row, i) {
        if (row[0] && String(row[0]).trim()) {
          appConfig[String(row[0]).trim()] = row[2] || '';
        }
      });
    }
  } catch (e) {
    appConfig = { error: e.toString() };
  }

  // クォータ情報
  let quotaInfo = {};
  try {
    quotaInfo = {
      remainingDailyQuota: MailApp.getRemainingDailyQuota()
    };
  } catch (e) {
    quotaInfo = { note: 'メールクォータ取得不可' };
  }

  return {
    success: true,
    timestamp: new Date().toISOString(),
    spreadsheet: spreadsheetInfo,
    sheets: sheetDetails,
    script: scriptInfo,
    triggers: triggers,
    moduleSettings: moduleSettings,
    appConfig: appConfig,
    quota: quotaInfo
  };
}
