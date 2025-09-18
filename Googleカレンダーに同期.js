function syncCalendars() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("年間行事予定表");
  const data = sheet.getDataRange().getValues();

  // 共通関数を使用してカレンダーIDを取得
  const 行事予定CalendarId = getOrCreateCalendarId('EVENT');
  const 対外行事CalendarId = getOrCreateCalendarId('EXTERNAL');

  Logger.log("[INFO] カレンダーID取得完了、同期処理を開始します。");

  const 行事予定Calendar = CalendarApp.getCalendarById(行事予定CalendarId);
  const 対外行事Calendar = CalendarApp.getCalendarById(対外行事CalendarId);
  const holidayCalendar = CalendarApp.getCalendarsByName('日本の祝日')[0];

  // カラム情報を定義
  const 行事予定Columns = [{ titleCol: 4 }]; // D列
  const 対外行事Columns = [{ titleCol: 13 }]; // M列

  // 行事予定と対外行事のイベント同期処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[1];

    if (!date) continue; // 日付が空の場合はスキップ

    processEventUpdates(sheet, 行事予定Calendar, 行事予定Columns, row, date, "行事予定", i + 1, holidayCalendar);
    processEventUpdates(sheet, 対外行事Calendar, 対外行事Columns, row, date, "対外行事", i + 1, holidayCalendar);
  }
  Logger.log("[INFO] カレンダーの同期が完了しました。");
}

function processEventUpdates(sheet, calendar, columns, row, date, eventType, rowIndex, holidayCalendar) {
  const holidays = holidayCalendar.getEventsForDay(date);
  const existingEvents = calendar.getEventsForDay(date) || [];
  let newEvents = [];
  let eventsChanged = false;

  columns.forEach(({ titleCol }) => {
    const titleCell = sheet.getRange(rowIndex, titleCol).getValue(); 
    if (titleCell) {
      const titles = titleCell
        .split('\n')
        .map(t => t.trim().replace(/^・/, '')) 
        .filter(t => t); 

      titles.forEach((title) => {
        const isHoliday = holidays.some(holiday => 
          holiday.getTitle() === title || 
          holiday.getTitle() === "振替休日" 
        );
        if (!isHoliday) {
          const eventInfo = parseEventTimesAndDates(title, date);
          if (eventInfo) {
            newEvents.push(eventInfo);
          }
        }
      });
    }
  });

  // 既存イベントをマッピング
  let existingEventMap = {};
  existingEvents.forEach(event => {
    const title = event.getTitle();
    const startTime = event.getStartTime().getTime();
    const endTime = event.getEndTime().getTime();
    const key = title + '_' + startTime + '_' + endTime;
    existingEventMap[key] = event;
  });

  // 新しいイベントをマッピング
  let newEventMap = {};
  newEvents.forEach(eventObj => {
    const title = eventObj.title;
    const startTime = eventObj.startTime.getTime();
    const endTime = eventObj.endTime.getTime();
    const key = title + '_' + startTime + '_' + endTime;
    newEventMap[key] = eventObj;
  });

  // 既存イベントの削除
  for (const key in existingEventMap) {
    if (!newEventMap[key]) {
      existingEventMap[key].deleteEvent();
      eventsChanged = true;
      Logger.log(`[INFO] 削除された${eventType}イベント: タイトル "${existingEventMap[key].getTitle()}"`);
    }
  }

  // 新しいイベントの作成
  for (const key in newEventMap) {
    if (!existingEventMap[key]) {
      const eventObj = newEventMap[key];
      if (eventObj.isAllDay) {
        calendar.createAllDayEvent(eventObj.title, eventObj.startTime, eventObj.endTime);
      } else {
        calendar.createEvent(eventObj.title, eventObj.startTime, eventObj.endTime);
      }
      eventsChanged = true;
      Logger.log(`[INFO] 新規作成された${eventType}イベント: タイトル "${eventObj.title}"、開始日時 ${eventObj.startTime}`);
    }
  }

  // 変更があった場合のみログ出力
  if (eventsChanged) {
    Logger.log(`[INFO] ${eventType}イベントの変更が完了しました。日付: ${date}`);
  }
}

function parseEventTimesAndDates(title, date) {
  let trimmedTitle = convertFullWidthToHalfWidth(title.trim());
  let originalTitle = trimmedTitle;

  let isAllDay = false;
  let startTime = new Date(date);
  let endTime = null;

  // 時間範囲のパターン（全角・半角両対応）
  const timePatternRange = /(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?(?:\s*[~～]\s*)(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?/;

  // 単一時間のパターン（全角・半角両対応）
  const timePatternSingle = /(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?/;

  // 時間範囲のマッチをチェック
  let timeMatch = trimmedTitle.match(timePatternRange);
  if (timeMatch) {
    let [_, startHour, startMinute, endHour, endMinute] = timeMatch;
    startTime = setEventTime(startTime, startHour, startMinute);
    endTime = setEventTime(new Date(date), endHour, endMinute);

    // 終了時間が開始時間より前の場合、翌日の時間として処理
    if (endTime <= startTime) {
      endTime.setDate(endTime.getDate() + 1);
    }

    isAllDay = false;
    return {startTime, endTime, title: originalTitle, isAllDay};
  }

  // 単一時間のマッチをチェック
  timeMatch = trimmedTitle.match(timePatternSingle);
  if (timeMatch) {
    let [_, startHour, startMinute] = timeMatch;
    startTime = setEventTime(startTime, startHour, startMinute);
    endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // 1時間後

    isAllDay = false;
    return {startTime, endTime, title: originalTitle, isAllDay};
  }

  // 時間情報が見つからない場合は全日イベントとして処理
  isAllDay = true;
  endTime = new Date(startTime.getTime() + 24 * 60 * 60 * 1000); // 終了日は含まれないため+1日
  return {startTime, endTime, title: originalTitle, isAllDay};
}

function setEventTime(date, hour, minute) {
  hour = parseInt(hour);
  minute = parseMinute(minute);
  date.setHours(hour, minute, 0, 0);
  return date;
}

function parseMinute(minute) {
  if (!minute || minute === '') return 0;
  if (minute === '半' || minute === '30分') return 30;
  return parseInt(minute.replace('分', '')) || 0;
}

function convertFullWidthToHalfWidth(str) {
  return str.replace(/[！-～]/g, function(tmpStr) {
    return String.fromCharCode(tmpStr.charCodeAt(0) - 0xFEE0);
  }).replace(/"/g, '"')
    .replace(/'/g, "'")
    .replace(/'/g, "`")
    .replace(/￥/g, "\\")
    .replace(/　/g, " ")
    .replace(/〜/g, "~");
}

function deleteAllEventsFromCalendars() {
  // 削除対象のカレンダーID
  var calendarIds = [
    'c_7b92329b928d730ba0860fffa81441c4348c4c1ad5110b411cbe3b4982259a1b@group.calendar.google.com',
    'c_2cc3d3fb5537385b1094a55da56d60832ea66eb03b186acc25d6ac5b40d409bf@group.calendar.google.com'
  ];

  // 各カレンダーの予定を全て削除
  for (var i = 0; i < calendarIds.length; i++) {
    var calendar = CalendarApp.getCalendarById(calendarIds[i]);
    var events = calendar.getEvents(new Date(2000, 0, 1), new Date(2100, 11, 31)); // 全てのイベント取得

    for (var j = 0; j < events.length; j++) {
      events[j].deleteEvent(); // イベント削除
    }
  }
  
  Logger.log('全ての予定が削除されました。');
}
