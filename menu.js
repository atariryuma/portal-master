/**
 * @fileoverview メニュー構成定義
 * @description ポータルマスターのメインメニューとサブメニューを構成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🎯 ポータルマスター');

  const introductionMenu = ui.createMenu('🚀 導入')
    .addItem('年間行事計画をインポート', 'importAnnualEvents')
    .addItem('行事予定表へ反映', 'updateAnnualEvents');

  const settingsMenu = ui.createMenu('⚙️ 設定')
    .addItem('年度更新設定', 'showAnnualUpdateSettingsDialog')
    .addItem('自動トリガー設定', 'showTriggerSettingsDialog');

  const dailyMenu = ui.createMenu('📅 日常業務')
    .addItem('今日の日付へ移動', 'setDailyHyperlink')
    .addItem('週報をPDF保存', 'saveToPDF')
    .addItem('週報フォルダを開く', 'openWeeklyReportFolder');

  const dutyMenu = ui.createMenu('👥 日直')
    .addItem('日直を自動割り当て', 'assignDuty')
    .addItem('日直のみ更新', 'updateAnnualDuty')
    .addItem('休業期間日直を集計', 'countStars');

  const reportMenu = ui.createMenu('📊 集計')
    .addItem('累計時数を計算', 'calculateCumulativeHours')
    .addItem('学年別授業時数を集計', 'aggregateSchoolEventsByGrade')
    .addItem('モジュール学習管理', 'showModulePlanningDialog');

  const integrationMenu = ui.createMenu('🔁 連携と年度更新')
    .addItem('カレンダーと同期', 'syncCalendars')
    .addItem('年度更新ファイル作成', 'copyAndClear');

  const helpMenu = ui.createMenu('❓ ヘルプ')
    .addItem('使い方ガイド', 'showUserGuide')
    .addItem('製作者情報', 'showCreatorInfo');

  menu.addSubMenu(introductionMenu)
    .addSubMenu(settingsMenu)
    .addSubMenu(dailyMenu)
    .addSubMenu(dutyMenu)
    .addSubMenu(reportMenu)
    .addSubMenu(integrationMenu)
    .addSubMenu(helpMenu)
    .addToUi();

  // 内部管理シートは通常利用で見せない（軽量な非表示のみ。完全初期化はモジュール機能の初回使用時に遅延実行）
  hideInternalSheetsForNormalUse_();
}

/**
 * 通常利用時にシートを非表示にする（内部ヘルパー）
 * @param {string} sheetName - 非表示にするシート名
 */
function hideSheetForNormalUse_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) {
    Logger.log('[WARNING] ' + sheetName + 'シートが見つかりません。');
    return;
  }
  if (targetSheet.isSheetHidden()) {
    return;
  }

  const visibleSheets = ss.getSheets().filter(function(sheet) {
    return !sheet.isSheetHidden();
  });
  if (visibleSheets.length <= 1) {
    Logger.log('[WARNING] 表示中シートが1枚のみのため、' + sheetName + 'シートを非表示にできません。');
    return;
  }

  const activeSheet = ss.getActiveSheet();
  if (activeSheet && activeSheet.getSheetId() === targetSheet.getSheetId()) {
    const fallbackSheet = visibleSheets.find(function(sheet) {
      return sheet.getSheetId() !== targetSheet.getSheetId();
    });
    if (fallbackSheet) {
      ss.setActiveSheet(fallbackSheet);
    }
  }

  targetSheet.hideSheet();
  Logger.log('[INFO] ' + sheetName + 'シートを非表示にしました。');
}

/**
 * 内部管理シートを非表示にする
 * @param {boolean=} includeMaster - マスターも含める場合true（テスト後クリーンアップ用）
 *   onOpenではfalse: マスターは初期セットアップ中にユーザーが編集するため隠さない。
 *   マスターの非表示は年間行事インポート完了時（updateAnnualEvents）で行う。
 */
function hideInternalSheetsForNormalUse_(includeMaster) {
  const sheetNames = [
    MODULE_SHEET_NAMES.CONTROL,
    SETTINGS_SHEET_NAME
  ];
  if (includeMaster) {
    sheetNames.unshift(MASTER_SHEET.NAME);
  }

  sheetNames.forEach(function(name) {
    try {
      hideSheetForNormalUse_(name);
    } catch (error) {
      Logger.log('[WARNING] ' + name + 'シートの非表示化に失敗: ' + error.toString());
    }
  });
}

function showCreatorInfo() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 10px;">
      <h3 style="color: #2c3e50; margin-top: 0;">製作者情報</h3>
      <p><strong>製作者:</strong> 中龍馬（Atari Ryuma）</p>
      <p><strong>連絡先:</strong> <a href="mailto:atarirym@open.ed.jp" style="color: #3498db;">atarirym@open.ed.jp</a></p>
      <hr style="border: none; border-top: 1px solid #ecf0f1; margin: 15px 0;">
      <p style="font-size: 0.9em; color: #7f8c8d;">
        このスプレッドシートは、学校業務の「テンプレート＋年間行事予定（評価語句）＋学年別時数」に
        いくつかの機能を追加したカスタムバージョンです。
      </p>
      <p style="font-size: 0.9em; color: #7f8c8d;">
        使用方法等について質問がある場合は、上記のメールアドレスまでご連絡ください。
      </p>
    </div>
  `)
    .setWidth(400)
    .setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ポータルマスター - 製作者情報');
}

/**
 * 使い方ガイドを表示する関数
 */
function showUserGuide() {
  try {
    const htmlFile = HtmlService.createHtmlOutputFromFile('userGuide');
    htmlFile.setWidth(1000).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(htmlFile, 'ポータルマスター - 使い方ガイド');
  } catch (error) {
    Logger.log('[WARNING] userGuide.html の読み込みに失敗しました: ' + error.toString());
    const fallbackHtml = HtmlService.createHtmlOutput(`
      <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 20px;">
        <h2 style="color: #2c3e50;">❓ ポータルマスター 使い方ガイド</h2>

        <h3>🚀 導入</h3>
        <ul>
          <li><strong>年間行事計画をインポート:</strong> Googleスプレッドシートから行事予定をインポート</li>
          <li><strong>行事予定表へ反映:</strong> マスターから年間行事予定表にデータを反映</li>
        </ul>

        <h3>⚙️ 設定</h3>
        <ul>
          <li><strong>年度更新設定:</strong> 年度更新・連携先ID・基準日を設定</li>
          <li><strong>自動トリガー設定:</strong> 自動処理のON/OFF・曜日・時刻を設定</li>
        </ul>

        <h3>📅 日常業務</h3>
        <ul>
          <li><strong>今日の日付へ移動:</strong> B1セルに今日の日付リンクを設定</li>
          <li><strong>週報をPDF保存:</strong> 週報シートをPDF形式で保存</li>
          <li><strong>週報フォルダを開く:</strong> PDF保存先フォルダをブラウザで開く</li>
        </ul>

        <h3>👥 日直</h3>
        <ul>
          <li><strong>日直を自動割り当て:</strong> 日直表を元にマスターへ日直を設定</li>
          <li><strong>日直のみ更新:</strong> マスターの日直列のみ年間行事予定表へ反映</li>
          <li><strong>休業期間日直を集計:</strong> 年間行事予定表の☆を日直ごとに集計</li>
        </ul>

        <h3>📊 集計</h3>
        <ul>
          <li><strong>累計時数を計算:</strong> 最新土曜日までの累計授業時数を更新</li>
          <li><strong>学年別授業時数を集計:</strong> 低中高学年別の詳細な時数レポート作成</li>
          <li><strong>モジュール学習管理:</strong> 計画・実施差分を管理して再集計</li>
        </ul>

        <h3>🔁 連携と年度更新</h3>
        <ul>
          <li><strong>カレンダーと同期:</strong> Googleカレンダーにイベントを同期します</li>
          <li><strong>年度更新ファイル作成:</strong> 新年度用にファイルをコピー・クリア</li>
        </ul>
      </div>
    `)
      .setWidth(800)
      .setHeight(600);

    SpreadsheetApp.getUi().showModalDialog(fallbackHtml, 'ポータルマスター - 使い方ガイド');
  }
}
