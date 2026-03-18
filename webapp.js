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
      if (e.parameter.format === 'json') {
        return buildJsonResponse_(runTestsViaWebapp_(e.parameter.suite || 'full', e.parameter.phase || ''));
      }
      return buildTestResultPage_(e.parameter.suite || 'quick', e.parameter.phase || '');
    case 'run':
      return buildRunResultPage_(e.parameter.fn || '');
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
  const data = webappGetDashboard();
  const html = buildDashboardHtml_(data);
  return HtmlService.createHtmlOutput(html)
    .setTitle('ポータルマスター')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ダッシュボードHTMLをサーバー側で組み立てる
 * @param {Object} data - webappGetDashboard()の結果
 * @return {string} HTML文字列
 */
function buildDashboardHtml_(data) {
  function esc(str) {
    return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // 今週の予定セクション
  let weekHtml = '';
  if (data.weekEvents && data.weekEvents.found) {
    weekHtml = '<table class="grade-table"><tr><th>日付</th><th>曜日</th><th>校内行事</th><th>日直</th></tr>';
    data.weekEvents.days.forEach(function(day) {
      const rowStyle = day.isToday ? ' style="background:#eef2ff;font-weight:bold;"' : '';
      const dateLabel = day.date.substring(5).replace('-', '/');
      const event = day.internalEvent || '-';
      const duty = day.duty || '-';
      const eventHtml = esc(event).replace(/\n/g, '<br>');
      const dutyHtml = esc(duty).replace(/\n/g, '<br>');
      weekHtml += '<tr' + rowStyle + '><td>' + esc(dateLabel) + '</td><td>' + esc(day.weekday) + '</td><td style="text-align:left;">' + eventHtml + '</td><td>' + dutyHtml + '</td></tr>';
    });
    weekHtml += '</table>';
  } else {
    weekHtml = '<div class="no-event">予定データなし</div>';
  }

  // 累計時数セクション
  let cumulativeHtml = '';
  if (data.cumulative && data.cumulative.found) {
    cumulativeHtml = '<div class="cumulative-header">' + esc(data.cumulative.header) + '</div>';
    cumulativeHtml += '<table class="grade-table"><tr><th>学年</th><th>授業時数</th></tr>';
    data.cumulative.grades.forEach(function(g) {
      cumulativeHtml += '<tr><td>' + g.grade + '年</td><td><strong>' + g.classHours + '</strong></td></tr>';
    });
    cumulativeHtml += '</table>';
  } else {
    cumulativeHtml = '<div class="no-event">累計時数データなし</div>';
  }

  // トリガー＋最終実行セクション
  let triggerHtml = '<table class="grade-table" style="font-size:0.82em;">';
  triggerHtml += '<tr><th>機能</th><th>状態</th><th>最終実行</th></tr>';
  const triggerFunctions = Object.keys(TRIGGER_FUNCTION_LABELS);
  const registeredSet = {};
  if (data.triggers) {
    data.triggers.forEach(function(t) { registeredSet[t['function']] = true; });
  }
  triggerFunctions.forEach(function(fn) {
    const label = TRIGGER_FUNCTION_LABELS[fn] || fn;
    const isRegistered = registeredSet[fn];
    const record = data.lastRun ? data.lastRun[fn] : null;

    // 状態アイコン
    let statusHtml = '';
    if (!isRegistered) {
      statusHtml = '<span style="color:#999;">未登録</span>';
    } else if (record && record.success) {
      statusHtml = '<span style="color:#4caf50;">&#9679; 正常</span>';
    } else if (record && !record.success) {
      statusHtml = '<span style="color:#d93025;">&#9679; 失敗</span>';
    } else {
      statusHtml = '<span style="color:#f29900;">&#9679; 未実行</span>';
    }

    // 最終実行日時
    let timeHtml = '-';
    if (record && record.timestamp) {
      const runDate = new Date(record.timestamp);
      timeHtml = Utilities.formatDate(runDate, 'Asia/Tokyo', 'MM/dd HH:mm');
      if (!record.success && record.error) {
        timeHtml += '<div style="color:#d93025;font-size:0.85em;">' + esc(record.error.substring(0, 40)) + '</div>';
      }
    }

    triggerHtml += '<tr><td style="text-align:left;">' + esc(label) + '</td><td>' + statusHtml + '</td><td>' + timeHtml + '</td></tr>';
  });
  triggerHtml += '</table>';

  // ヘルスセクション
  const h = data.health || { sheetsOk: 0, sheetsMissing: 0, totalSheets: 9, details: [] };
  const pct = h.totalSheets > 0 ? Math.round((h.sheetsOk / h.totalSheets) * 100) : 0;
  const barClass = pct === 100 ? 'health-ok' : 'health-warn';
  let healthHtml = '<div class="health-bar">';
  healthHtml += '<div class="health-fill"><div class="health-fill-inner ' + barClass + '" style="width:' + pct + '%"></div></div>';
  healthHtml += '<span class="health-text">' + h.sheetsOk + '/' + h.totalSheets + ' OK</span></div>';
  // シート一覧
  if (h.details && h.details.length > 0) {
    healthHtml += '<div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:4px;">';
    h.details.forEach(function(s) {
      const bg = s.ok ? '#e8f5e9' : '#fce4ec';
      const color = s.ok ? '#2e7d32' : '#c62828';
      const icon = s.ok ? '&#9679;' : '&#10005;';
      healthHtml += '<span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:0.78em;background:' + bg + ';color:' + color + ';">' + icon + ' ' + esc(s.name) + '</span>';
    });
    healthHtml += '</div>';
  }

  // 診断セクション
  let diagnosticsHtml = '';
  if (data.diagnosticChecks) {
    const allOk = data.diagnosticChecks.every(function(c) { return c.ok; });
    const failedChecks = data.diagnosticChecks.filter(function(c) { return !c.ok; });

    if (allOk) {
      diagnosticsHtml = '<div style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:#e8f5e9;border-radius:6px;font-size:0.85em;">'
        + '<span style="color:#4caf50;font-size:1.2em;">&#9679;</span>'
        + '関数チェック: ' + data.diagnosticChecks.length + '/' + data.diagnosticChecks.length + ' 正常'
        + '</div>';
    } else {
      diagnosticsHtml = '<div style="padding:8px 12px;background:#fce4ec;border-radius:6px;font-size:0.85em;">'
        + '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">'
        + '<span style="color:#d93025;font-size:1.2em;">&#9679;</span>'
        + '関数チェック: ' + (data.diagnosticChecks.length - failedChecks.length) + '/' + data.diagnosticChecks.length + ' 正常'
        + '</div>';
      failedChecks.forEach(function(c) {
        diagnosticsHtml += '<div style="color:#d93025;font-size:0.8em;margin-left:20px;">' + esc(c.name) + ': ' + esc(c.error.substring(0, 60)) + '</div>';
      });
      diagnosticsHtml += '</div>';
    }
  }

  // ヘッダー情報
  const ts = new Date(data.timestamp);
  const timeStr = Utilities.formatDate(ts, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const ssName = data.spreadsheet ? esc(data.spreadsheet.name) : '';
  const ssUrl = data.spreadsheet ? esc(data.spreadsheet.url) : '';
  const baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:"Yu Gothic","Hiragino Sans",Arial,sans-serif;background:#f0f2f5;color:#1a1a2e;min-height:100vh;padding:20px}'
    + '.dashboard{max-width:800px;margin:0 auto}'
    + '.header{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;border-radius:12px;padding:24px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}'
    + '.header h1{font-size:1.4em}.header-sub{font-size:0.85em;opacity:0.85;margin-top:4px}'
    + '.header-right{text-align:right;font-size:0.8em;opacity:0.8}'
    + '.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}'
    + '@media(max-width:600px){.grid{grid-template-columns:1fr}}'
    + '.card{background:#fff;border-radius:10px;padding:18px;box-shadow:0 1px 4px rgba(0,0,0,0.06)}'
    + '.card h2{font-size:0.9em;color:#667eea;margin-bottom:10px;display:flex;align-items:center;gap:6px}'
    + '.event-main{font-size:1.2em;font-weight:bold;color:#2c3e50;padding:8px 0;border-bottom:1px solid #f0f0f0;margin-bottom:8px}'
    + '.event-row{display:flex;justify-content:space-between;padding:4px 0;font-size:0.88em}'
    + '.event-label{color:#888}.event-value{font-weight:600}'
    + '.grade-table{width:100%;border-collapse:collapse;font-size:0.85em}'
    + '.grade-table th{background:#f8f9fa;padding:6px 8px;text-align:center;font-weight:600;color:#555;border-bottom:2px solid #e0e0e0}'
    + '.grade-table td{padding:6px 8px;text-align:center;border-bottom:1px solid #f0f0f0}'
    + '.grade-table tr:hover{background:#f8f9ff}'
    + '.trigger-list{list-style:none}'
    + '.trigger-item{display:flex;align-items:center;gap:8px;padding:5px 0;font-size:0.85em;border-bottom:1px solid #f8f8f8}'
    + '.trigger-item:last-child{border-bottom:none}'
    + '.dot{width:8px;height:8px;border-radius:50%;background:#4caf50;flex-shrink:0}'
    + '.health-bar{display:flex;align-items:center;gap:10px;margin-top:8px}'
    + '.health-fill{flex:1;height:8px;background:#e0e0e0;border-radius:4px;overflow:hidden}'
    + '.health-fill-inner{height:100%;border-radius:4px}'
    + '.health-ok{background:#4caf50}.health-warn{background:#f29900}'
    + '.health-text{font-size:0.85em;font-weight:600;white-space:nowrap}'
    + '.no-event{color:#999;font-style:italic;font-size:0.9em;padding:8px 0}'
    + '.cumulative-header{font-size:0.8em;color:#888;margin-bottom:6px}'
    + '.ss-link{color:#fff;text-decoration:none;opacity:0.9;font-size:0.85em}'
    + '.ss-link:hover{opacity:1;text-decoration:underline}'
    + '.card-full{grid-column:1/-1}'
    + '.actions{display:flex;gap:8px;flex-wrap:wrap}'
    + '.btn-action{display:inline-block;padding:10px 16px;border-radius:6px;text-decoration:none;font-size:0.85em;font-weight:600;color:#fff;text-align:center;flex:1;min-width:120px;transition:all 0.2s}'
    + '.btn-action:hover{opacity:0.85}'
    + '.btn-action.running{opacity:0.6;pointer-events:none}'
    + '@keyframes spin{to{transform:rotate(360deg)}}'
    + '.spinner{display:inline-block;width:14px;height:14px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spin 0.8s linear infinite;vertical-align:middle;margin-right:6px}'
    + '.footer{margin-top:20px;text-align:center;color:#bdc3c7;font-size:0.75em}'
    + '</style></head><body>'
    + '<div class="dashboard">'
    + '<div class="header"><div><h1>ポータルマスター</h1>'
    + '<div class="header-sub"><a class="ss-link" href="' + ssUrl + '" target="_blank">' + ssName + ' &#8599;</a></div>'
    + '</div><div class="header-right">' + esc(timeStr) + '</div></div>'
    + '<div class="grid">'
    + '<div class="card card-full"><h2>&#128197; 今週の予定</h2>' + weekHtml + '</div>'
    + '<div class="card"><h2>&#128202; 累計授業時数</h2>' + cumulativeHtml + '</div>'
    + '<div class="card"><h2>&#9200; トリガー・実行状況</h2>' + triggerHtml + '</div>'
    + '<div class="card"><h2>&#9989; システム状態</h2>' + healthHtml + '</div>'
    + '<div class="card"><h2>&#9889; クイック操作</h2>'
    + '<div class="actions">'
    + '<a class="btn-action" style="background:#27ae60;" href="' + esc(baseUrl) + '?page=run&fn=calculateCumulativeHours" onclick="return runWithFeedback(this,\'累計時数を再計算しますか？\')">累計時数を再計算</a>'
    + '<a class="btn-action" style="background:#e67e22;" href="' + esc(baseUrl) + '?page=run&fn=saveToPDF" onclick="return runWithFeedback(this,\'週報をPDF保存しますか？\')">週報PDF保存</a>'
    + '<a class="btn-action" style="background:#8e44ad;" href="' + esc(baseUrl) + '?page=run&fn=syncCalendars" onclick="return runWithFeedback(this,\'カレンダーと同期しますか？\')">カレンダー同期</a>'
    + '</div></div>'
    + '</div>'
    + '<div class="card" style="margin-top:12px;"><h2>&#128295; メンテナンス</h2>'
    + diagnosticsHtml
    + '<div class="actions" style="margin-top:12px;">'
    + '<a class="btn-action" style="background:#34495e;" href="' + esc(baseUrl) + '?page=run&fn=setDailyHyperlink" onclick="return runWithFeedback(this)">日付リンク更新</a>'
    + '<a class="btn-action" style="background:#2c3e50;" href="' + esc(baseUrl) + '?page=tests&suite=quick" onclick="return runWithFeedback(this)">クイックテスト</a>'
    + '<a class="btn-action" style="background:#1a252f;" href="' + esc(baseUrl) + '?page=tests&suite=full&phase=8" onclick="return runWithFeedback(this)">追加テスト(Phase8)</a>'
    + '</div></div>'
    + '<div class="footer">Portal Master Dashboard</div>'
    + '</div>'
    + '<script>'
    + 'function runWithFeedback(el,msg){'
    + 'if(msg&&!confirm(msg))return false;'
    + 'var label=el.textContent;'
    + 'el.classList.add("running");'
    + 'el.innerHTML="<span class=spinner></span>"+label+"...";'
    + 'return true;'
    + '}'
    + '</script>'
    + '</body></html>';
}

// ========================================
// ダッシュボード用公開ラッパー
// google.script.run は末尾 _ の関数を呼べないため公開名で中継する
// ========================================

function webappGetStatus() {
  return getStatusInfo_();
}

function webappRunDiagnostics() {
  return runDiagnostics_();
}

function webappRunTests(suite, phase) {
  return runTestsViaWebapp_(suite || 'quick', phase || '');
}

/**
 * ダッシュボード用データを一括取得する
 * @return {Object} ダッシュボードに必要な全データ
 */
function webappGetDashboard() {
  let ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // ウェブアプリコンテキストではアクティブなスプレッドシートがない場合がある
  }
  if (!ss) {
    // バインドスクリプトの場合、PropertiesServiceからIDを取得するか、直接指定
    return { error: 'スプレッドシートを取得できません。ウェブアプリのコンテキストを確認してください。', timestamp: new Date().toISOString() };
  }
  const now = new Date();

  // 今週の行事予定を取得
  const weekEvents = getWeekEvents_(ss, now);

  // 累計時数を取得
  const cumulative = getCumulativeData_(ss);

  // 最終実行記録
  const lastRun = getLastRunRecords_();

  // 診断結果（軽量：関数チェックのみ）
  const diagnosticChecks = [];
  const checkFns = [
    { name: 'getSettingsSheetOrThrow', fn: function() { getSettingsSheetOrThrow(); } },
    { name: 'readModuleSettingsMap', fn: function() { readModuleSettingsMap(); } },
    { name: 'getCurrentOrNextSaturday', fn: function() { getCurrentOrNextSaturday(); } },
    { name: 'getFiscalYear', fn: function() { getFiscalYear(new Date()); } }
  ];
  checkFns.forEach(function(check) {
    try {
      check.fn();
      diagnosticChecks.push({ name: check.name, ok: true });
    } catch (error) {
      diagnosticChecks.push({ name: check.name, ok: false, error: error.toString() });
    }
  });

  // トリガー状態
  const triggers = ScriptApp.getProjectTriggers().map(function(t) {
    return {
      function: t.getHandlerFunction(),
      type: t.getEventType().toString()
    };
  });

  // 診断サマリー（シートごとの状態）
  let sheetOk = 0;
  let sheetMissing = 0;
  const sheetChecks = [];
  REQUIRED_SHEET_NAMES.forEach(function(name) {
    const exists = !!ss.getSheetByName(name);
    if (exists) { sheetOk++; } else { sheetMissing++; }
    sheetChecks.push({ name: name, ok: exists });
  });

  return {
    timestamp: now.toISOString(),
    spreadsheet: {
      name: ss.getName(),
      url: ss.getUrl()
    },
    weekEvents: weekEvents,
    cumulative: cumulative,
    lastRun: lastRun,
    diagnosticChecks: diagnosticChecks,
    triggers: triggers,
    health: {
      sheetsOk: sheetOk,
      sheetsMissing: sheetMissing,
      totalSheets: REQUIRED_SHEET_NAMES.length,
      details: sheetChecks
    }
  };
}

/**
 * 今週の行事予定を年間行事予定表から取得する（今日〜5日先）
 * @param {Spreadsheet} ss
 * @param {Date} today
 * @return {Object} 今週のイベント情報
 */
function getWeekEvents_(ss, today) {
  const sheet = ss.getSheetByName(ANNUAL_SCHEDULE.SHEET_NAME);
  if (!sheet) {
    return { found: false, days: [], message: '年間行事予定表が見つかりません' };
  }

  // 今日から5日分の日付キーを生成
  const targetDates = {};
  for (let d = 0; d < 5; d++) {
    const date = new Date(today.getFullYear(), today.getMonth(), today.getDate() + d);
    targetDates[formatDateKey(date)] = { offset: d, date: date };
  }

  // 日付列とイベントデータを1回のバッチ読み取りで取得
  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(ANNUAL_SCHEDULE.DATA_START_ROW, 1, lastRow - ANNUAL_SCHEDULE.DATA_START_ROW + 1, ANNUAL_SCHEDULE.DUTY_COLUMN).getValues();

  const dateColIdx = ANNUAL_SCHEDULE.DATE_COLUMN.charCodeAt(0) - 'A'.charCodeAt(0);
  const weekdayNames = ['日', '月', '火', '水', '木', '金', '土'];
  const days = [];
  const foundKeys = {};

  for (let i = 0; i < allData.length; i++) {
    const cellDate = allData[i][dateColIdx];
    if (cellDate instanceof Date) {
      const key = formatDateKey(cellDate);
      if (targetDates[key] && !foundKeys[key]) {
        foundKeys[key] = true;
        const info = targetDates[key];
        days.push({
          date: key,
          weekday: weekdayNames[info.date.getDay()],
          isToday: info.offset === 0,
          internalEvent: String(allData[i][ANNUAL_SCHEDULE.INTERNAL_EVENT_COLUMN - 1] || ''),
          externalEvent: String(allData[i][ANNUAL_SCHEDULE.EXTERNAL_EVENT_COLUMN - 1] || ''),
          duty: String(allData[i][ANNUAL_SCHEDULE.DUTY_COLUMN - 1] || '')
        });
      }
    }
  }
  days.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  return { found: days.length > 0, days: days };
}

/**
 * ワンタップ操作の実行結果ページを生成する
 * @param {string} fnName - 実行する関数名
 * @return {HtmlOutput} 結果ページ
 */
function buildRunResultPage_(fnName) {
  const allowedFunctions = {
    calculateCumulativeHours: '累計時数を計算',
    syncCalendars: 'カレンダーと同期',
    setDailyHyperlink: '今日の日付へ移動',
    saveToPDF: '週報をPDF保存'
  };

  const label = allowedFunctions[fnName];
  if (!label) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center;">'
      + '<h2 style="color:#d93025;">不正な操作です</h2>'
      + '<p><a href="' + ScriptApp.getService().getUrl() + '">ダッシュボードに戻る</a></p>'
      + '</body></html>'
    ).setTitle('エラー');
  }

  let resultMsg = '';
  let success = true;
  try {
    if (fnName === 'calculateCumulativeHours') {
      calculateCumulativeHours();
      resultMsg = '累計時数の計算が完了しました。';
    } else if (fnName === 'syncCalendars') {
      syncCalendars();
      resultMsg = 'カレンダー同期が完了しました。';
    } else if (fnName === 'setDailyHyperlink') {
      setDailyHyperlink();
      resultMsg = '今日の日付リンクを設定しました。';
    } else if (fnName === 'saveToPDF') {
      saveToPDF();
      resultMsg = '週報のPDF保存が完了しました。';
    }
  } catch (e) {
    success = false;
    resultMsg = 'エラー: ' + e.toString();
  }

  const color = success ? '#27ae60' : '#d93025';
  const icon = success ? '&#9989;' : '&#10060;';
  const baseUrl = ScriptApp.getService().getUrl();

  // 他の操作リンク（現在実行した操作は除外）
  let otherActions = '';
  Object.keys(allowedFunctions).forEach(function(fn) {
    if (fn !== fnName) {
      const colors = { calculateCumulativeHours: '#27ae60', saveToPDF: '#e67e22', syncCalendars: '#8e44ad', setDailyHyperlink: '#34495e' };
      otherActions += '<a href="' + baseUrl + '?page=run&fn=' + fn + '" style="display:inline-block;background:' + (colors[fn] || '#555') + ';color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none;font-size:0.85em;font-weight:600;">' + allowedFunctions[fn] + '</a> ';
    }
  });

  return HtmlService.createHtmlOutput(
    '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="font-family:Yu Gothic,sans-serif;padding:20px;text-align:center;background:#f0f2f5;">'
    + '<div style="max-width:500px;margin:0 auto;background:#fff;border-radius:12px;padding:32px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">'
    + '<div style="font-size:3em;margin-bottom:16px;">' + icon + '</div>'
    + '<h2 style="color:' + color + ';margin-bottom:12px;">' + label + '</h2>'
    + '<p style="color:#555;margin-bottom:20px;">' + resultMsg + '</p>'
    + '<a href="' + baseUrl + '" style="display:inline-block;background:#667eea;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:bold;">ダッシュボードに戻る</a>'
    + '<div style="margin-top:20px;padding-top:16px;border-top:1px solid #eee;">'
    + '<div style="font-size:0.8em;color:#999;margin-bottom:8px;">他の操作</div>'
    + '<div style="display:flex;flex-wrap:wrap;gap:6px;justify-content:center;">' + otherActions + '</div>'
    + '</div>'
    + '</div></body></html>'
  ).setTitle(label);
}

/**
 * テスト結果をHTMLページで表示する
 * @param {string} suite - 'full' または 'quick'
 * @param {string} phase - フェーズ番号
 * @return {HtmlOutput} テスト結果ページ
 */
function buildTestResultPage_(suite, phase) {
  const data = runTestsViaWebapp_(suite, phase);
  const baseUrl = ScriptApp.getService().getUrl();
  const s = data.summary;
  const allPass = s.failed === 0;
  const icon = allPass ? '&#9989;' : '&#9888;';
  const color = allPass ? '#27ae60' : '#e67e22';
  const title = (suite === 'quick' ? 'クイック' : 'フル') + 'テスト' + (phase ? ' Phase' + phase : '');

  let html = '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="font-family:Yu Gothic,sans-serif;padding:20px;background:#f0f2f5;">'
    + '<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">'
    + '<div style="text-align:center;margin-bottom:16px;">'
    + '<div style="font-size:2.5em;">' + icon + '</div>'
    + '<h2 style="color:' + color + ';">' + title + '</h2>'
    + '<p style="font-size:1.2em;margin:8px 0;">' + s.passed + '/' + s.total + ' PASS (' + s.successRate + ')</p>'
    + '</div>';

  // テスト結果テーブル
  html += '<table style="width:100%;border-collapse:collapse;font-size:0.82em;">';
  html += '<tr style="background:#f8f9fa;"><th style="padding:6px 8px;text-align:left;border-bottom:2px solid #e0e0e0;">テスト</th><th style="padding:6px 8px;text-align:center;border-bottom:2px solid #e0e0e0;">結果</th></tr>';

  data.details.forEach(function(d) {
    let statusIcon = '';
    let statusColor = '';
    if (d.status === 'PASS') {
      statusIcon = '&#9679;';
      statusColor = '#4caf50';
    } else if (d.status === 'FAIL' || d.status === 'ERROR') {
      statusIcon = '&#9679;';
      statusColor = '#d93025';
    } else {
      statusIcon = '&#9679;';
      statusColor = '#f29900';
    }
    html += '<tr style="border-bottom:1px solid #f0f0f0;">';
    html += '<td style="padding:5px 8px;">' + d.name + '</td>';
    html += '<td style="padding:5px 8px;text-align:center;"><span style="color:' + statusColor + ';">' + statusIcon + ' ' + d.status + '</span></td>';
    html += '</tr>';
    if (d.status !== 'PASS' && d.message) {
      html += '<tr><td colspan="2" style="padding:2px 8px 6px 24px;color:#d93025;font-size:0.9em;">' + d.message + '</td></tr>';
    }
  });
  html += '</table>';

  // エラーサマリー
  if (data.errors.length > 0) {
    html += '<div style="background:#fce4ec;border-radius:6px;padding:10px;margin-top:12px;font-size:0.8em;">';
    html += '<strong style="color:#d93025;">エラー詳細:</strong>';
    data.errors.forEach(function(e) {
      html += '<div style="margin-top:4px;color:#c62828;">' + e + '</div>';
    });
    html += '</div>';
  }

  html += '<div style="text-align:center;margin-top:16px;">'
    + '<a href="' + baseUrl + '" style="display:inline-block;background:#667eea;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:bold;">ダッシュボードに戻る</a>'
    + '</div>'
    + '<div style="margin-top:16px;padding-top:12px;border-top:1px solid #eee;display:flex;flex-wrap:wrap;gap:6px;justify-content:center;">'
    + '<a href="' + baseUrl + '?page=tests&suite=quick" style="display:inline-block;background:#2c3e50;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none;font-size:0.85em;font-weight:600;">クイックテスト再実行</a>'
    + '<a href="' + baseUrl + '?page=tests&suite=full&phase=8" style="display:inline-block;background:#1a252f;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none;font-size:0.85em;font-weight:600;">Phase8テスト</a>'
    + '</div>'
    + '</div></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

/**
 * 累計時数データを取得する
 * @param {Spreadsheet} ss
 * @return {Object} 累計時数情報
 */
function getCumulativeData_(ss) {
  const sheet = ss.getSheetByName(CUMULATIVE_SHEET.NAME);
  if (!sheet) {
    return { found: false, message: '累計時数シートが見つかりません' };
  }

  const header = sheet.getRange(CUMULATIVE_SHEET.DATE_CELL).getDisplayValue();
  const gradeData = sheet.getRange(CUMULATIVE_SHEET.GRADE_START_ROW, 3, 6, 1).getValues();
  const grades = [];
  for (let g = 0; g < 6; g++) {
    grades.push({
      grade: g + 1,
      classHours: Number(gradeData[g][0]) || 0
    });
  }

  return {
    found: true,
    header: header,
    grades: grades
  };
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
  REQUIRED_SHEET_NAMES.forEach(function(name) {
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
