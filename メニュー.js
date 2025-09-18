function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🎯 ポータルマスター');
  
  // 🏗️ 初回セットアップのサブメニュー
  var setupMenu = ui.createMenu('🏗️ 初回セットアップ')
      .addItem('📥 年間行事計画をインポート', 'importAnnualEvents')
      .addItem('👥 日直を自動割り当て', 'assignDuty')
      .addItem('📋 行事予定表へ反映', 'updateAnnualEvents')
      .addItem('📊 累計時数を初期計算', 'calculateCumulativeHours')
      .addItem('⚙️ 自動処理を設定', 'setAutomaticProcesses');
  
  // 📅 日常業務のサブメニュー
  var dailyMenu = ui.createMenu('📅 日常業務')
      .addItem('📄 週報をPDF保存', 'saveToPDF')
      .addItem('📁 週報フォルダを開く', 'openWeeklyReportFolder')
      .addItem('⭐ 休業期間日直をカウント', 'countStars')
      .addItem('🔗 今日の日付へ移動', 'setDailyHyperlink');
  
  // 📊 集計・レポートのサブメニュー
  var reportMenu = ui.createMenu('📊 集計・レポート')
      .addItem('📈 学年別授業時数を集計', 'aggregateSchoolEventsByGrade')
      .addItem('🧮 累計時数を再計算', 'calculateCumulativeHours')
      .addItem('📋 日直のみ更新', 'updateAnnualDuty');
  
  // 🔧 システム管理のサブメニュー
  var systemMenu = ui.createMenu('🔧 システム管理')
      .addItem('📅 カレンダーと同期', 'syncCalendars')
      .addItem('📋 年度更新ファイル作成', 'CopyAndClear')
      .addItem('❓ 使い方ガイドを表示', 'showUserGuide');
  
  // メインメニューにサブメニューを追加
  menu.addSubMenu(setupMenu)
      .addSubMenu(dailyMenu)
      .addSubMenu(reportMenu)
      .addSubMenu(systemMenu);
  
  // セパレーターと製作者情報を追加
  menu.addSeparator()
      .addItem('ℹ️ 製作者情報', 'showCreatorInfo');
  
  // メニューをUIに追加
  menu.addToUi();
}

function showCreatorInfo() {
  var htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 10px;">
      <h3 style="color: #2c3e50; margin-top: 0;">ℹ️ 製作者情報</h3>
      <p><strong>製作者:</strong> 中龍馬（Atari Ryuma）</p>
      <p><strong>連絡先:</strong> <a href="mailto:atarirym@open.ed.jp" style="color: #3498db;">atarirym@open.ed.jp</a></p>
      <hr style="border: none; border-top: 1px solid #ecf0f1; margin: 15px 0;">
      <p style="font-size: 0.9em; color: #7f8c8d;">
        ※このスプレッドシートは、研究所配布の「【テンプレート】年間行事予定表（編集用）小学校」に
        いくつかの機能を追加したカスタムバージョンです。
      </p>
      <p style="font-size: 0.9em; color: #7f8c8d;">
        使用方法等について質問がある場合は、上記のメールアドレス、または勤務校まで電話してください。
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
    // HTMLファイルからガイドを読み込んで表示
    var htmlFile = HtmlService.createHtmlOutputFromFile('ポータルマスター使い方ガイド');
    htmlFile.setWidth(1000).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(htmlFile, 'ポータルマスター - 使い方ガイド');
  } catch (error) {
    // HTMLファイルが見つからない場合の代替処理
    var fallbackHtml = HtmlService.createHtmlOutput(`
      <div style="font-family: 'Yu Gothic', Arial, sans-serif; padding: 20px;">
        <h2 style="color: #2c3e50;">❓ ポータルマスター 使い方ガイド</h2>
        <h3>🏗️ 初回セットアップ（順番通りに実行）</h3>
        <ol>
          <li><strong>年間行事計画をインポート:</strong> Excelファイルから行事予定をインポート</li>
          <li><strong>日直を自動割り当て:</strong> 日直表を基にマスターシートに日直を割り当て</li>
          <li><strong>行事予定表へ反映:</strong> マスターから年間行事予定表にデータを反映</li>
          <li><strong>累計時数を初期計算:</strong> 現在までの累計授業時数を計算</li>
          <li><strong>自動処理を設定:</strong> 定期実行トリガーを設定</li>
        </ol>
        
        <h3>📅 日常業務</h3>
        <ul>
          <li><strong>週報をPDF保存:</strong> 週報シートをPDF形式で保存</li>
          <li><strong>週報フォルダを開く:</strong> PDF保存先フォルダをブラウザで開く</li>
          <li><strong>休業期間日直をカウント:</strong> ☆マークの日直回数を集計</li>
          <li><strong>今日の日付へ移動:</strong> B1セルに今日の日付リンクを設定</li>
        </ul>
        
        <h3>📊 集計・レポート</h3>
        <ul>
          <li><strong>学年別授業時数を集計:</strong> 低中高学年別の詳細な時数レポート作成</li>
          <li><strong>累計時数を再計算:</strong> 土曜日までの累計時数を更新</li>
          <li><strong>日直のみ更新:</strong> 日直情報のみを部分更新</li>
        </ul>
        
        <h3>🔧 システム管理</h3>
        <ul>
          <li><strong>カレンダーと同期:</strong> GoogleカレンダーにイベントをProfessionalSync</li>
          <li><strong>年度更新ファイル作成:</strong> 新年度用にファイルをコピー・クリア</li>
        </ul>
        
        <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin-top: 20px;">
          <h4 style="color: #856404; margin: 0 0 10px 0;">⚠️ 重要な注意点</h4>
          <ul style="margin: 0;">
            <li>初回セットアップは必ず順番通りに実行してください</li>
            <li>「年度更新作業」シートに必要な設定値を事前に入力してください</li>
            <li>エラーが発生した場合は、スクリプトエディタでログを確認してください</li>
          </ul>
        </div>
      </div>
    `)
    .setWidth(800)
    .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(fallbackHtml, 'ポータルマスター - 使い方ガイド');
  }
}
