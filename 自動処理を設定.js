function setAutomaticProcesses() {
  createWeeklyTrigger();
  createDailyTrigger();
  createDailyTrigger2();
  createTrigger();
  Logger.log('自動処理のトリガーが設定されました。');

  // ダイアログを表示する部分を追加
  showDialog();
}

function createWeeklyTrigger() {
  ScriptApp.newTrigger('saveToPDF')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();
}

function createDailyTrigger() {
  ScriptApp.newTrigger('setDailyHyperlink')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();
}

function createDailyTrigger2() {
  ScriptApp.newTrigger('syncCalendars')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function createTrigger() {
  ScriptApp.newTrigger('calculateCumulativeHours')
    .timeBased()
    .atHour(2)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .create();
}

function showDialog() {
  var ui = SpreadsheetApp.getUi(); // スプレッドシートのUIを取得
  var message = '自動処理のトリガーが設定されました。\n\n' +
                '以下の処理が自動で実行されます：\n' +
                '1. 毎週月曜日の午前2時に「週報をPDFで保存」関数が実行されます。\n' +
                '2. 毎週月曜日の午前2時に「累計時数計算」関数が実行されます。\n' +
                '3. 毎日午前3時に「Googleカレンダーと同期」関数が実行されます。\n' +
                '4. 毎日午前4時に「今日の日付へ移動」関数が実行されます。\n\n' +
                '今後はこれらのスクリプトを手動で実行する必要はありません。';
  ui.alert('設定完了', message, ui.ButtonSet.OK);
}
