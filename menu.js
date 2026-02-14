/**
 * @fileoverview メニュー構成定義
 * @description ポータルマスターのメインメニューとサブメニューを構成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🎯 ポータルマスター');

  const setupMenu = ui.createMenu('🏗️ 初回セットアップ')
    .addItem('ステップ1: 年間行事計画をインポート', 'importAnnualEvents')
    .addItem('ステップ2: 日直を自動割り当て', 'assignDuty')
    .addItem('ステップ3: 行事予定表へ反映', 'updateAnnualEvents')
    .addItem('ステップ4: 累計時数を初期計算', 'calculateCumulativeHours')
    .addItem('ステップ5: 自動処理を設定', 'showTriggerSettingsDialog');

  const dailyMenu = ui.createMenu('📅 日常業務')
    .addItem('📄 週報をPDF保存', 'saveToPDF')
    .addItem('📁 週報フォルダを開く', 'openWeeklyReportFolder')
    .addItem('⭐ 休業期間日直をカウント', 'countStars')
    .addItem('🔗 今日の日付へ移動', 'setDailyHyperlink');

  const reportMenu = ui.createMenu('📊 集計・レポート')
    .addItem('📈 学年別授業時数を集計', 'aggregateSchoolEventsByGrade')
    .addItem('📋 日直のみ更新', 'updateAnnualDuty');

  const systemMenu = ui.createMenu('🔧 システム管理')
    .addItem('⚙️ 自動トリガー設定', 'showTriggerSettingsDialog')
    .addItem('🧩 モジュール学習管理', 'showModulePlanningDialog')
    .addItem('📅 カレンダーと同期', 'syncCalendars')
    .addItem('📋 年度更新ファイル作成', 'copyAndClear');

  menu.addItem('❓ 使い方ガイド', 'showUserGuide')
    .addSeparator()
    .addSubMenu(setupMenu)
    .addSubMenu(dailyMenu)
    .addSubMenu(reportMenu)
    .addSubMenu(systemMenu)
    .addSeparator()
    .addItem('ℹ️ 製作者情報', 'showCreatorInfo')
    .addToUi();
}

function showCreatorInfo() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 10px;">
      <h3 style="color: #2c3e50; margin-top: 0;">製作者情報</h3>
      <p><strong>製作者:</strong> 当たり竜馬 (Atari Ryuma)</p>
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
    const fallbackHtml = HtmlService.createHtmlOutput(`
      <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 20px;">
        <h2 style="color: #2c3e50;">❓ ポータルマスター 使い方ガイド</h2>
        <h3>🏗️ 初回セットアップ（順番通りに実行）</h3>
        <ol>
          <li><strong>年間行事計画をインポート:</strong> Googleスプレッドシートから行事予定をインポート</li>
          <li><strong>日直を自動割り当て:</strong> 日直表を元にマスターへ日直を設定</li>
          <li><strong>行事予定表へ反映:</strong> マスターから年間行事予定表にデータを反映</li>
          <li><strong>累計時数を初期計算:</strong> 現在までの累計授業時数を計算</li>
          <li><strong>自動処理を設定:</strong> 定期実行トリガーを設定</li>
        </ol>

        <h3>📅 日常業務</h3>
        <ul>
          <li><strong>週報をPDF保存:</strong> 週報シートをPDF形式で保存</li>
          <li><strong>週報フォルダを開く:</strong> PDF保存先フォルダをブラウザで開く</li>
          <li><strong>休業期間日直をカウント:</strong> 年間行事予定表の☆を日直ごとに集計</li>
          <li><strong>今日の日付へ移動:</strong> B1セルに今日の日付リンクを設定</li>
        </ul>

        <h3>📊 集計・レポート</h3>
        <ul>
          <li><strong>学年別授業時数を集計:</strong> 低中高学年別の詳細な時数レポート作成</li>
          <li><strong>日直のみ更新:</strong> マスターの日直列のみ年間行事予定表へ反映</li>
        </ul>

        <h3>🔧 システム管理</h3>
        <ul>
          <li><strong>モジュール学習管理:</strong> module_cycle_plan（2か月クール）を基準に日次計画を自動生成</li>
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
