/**
 * @fileoverview Googleカレンダー同期機能
 * @description 年間行事予定表の内容をGoogleカレンダーに同期し、
 *              校内行事・対外行事を別々のカレンダーに登録します。
 *              時間指定、全日イベント、祝日の自動スキップに対応。
 */

const CALENDAR_SYNC_MANAGED_MARKER = '[PORTAL_MASTER_MANAGED]';

function syncCalendars() {
  const sheet = getAnnualScheduleSheetOrThrow(); // 共通関数を使用してエラーハンドリング
  const data = sheet.getDataRange().getValues();

  // 共通関数を使用してカレンダーIDを取得
  const eventCalendarId = getOrCreateCalendarId('EVENT');
  const externalCalendarId = getOrCreateCalendarId('EXTERNAL');

  Logger.log("[INFO] カレンダーID取得完了、同期処理を開始します。");

  const eventCalendar = CalendarApp.getCalendarById(eventCalendarId);
  const externalCalendar = CalendarApp.getCalendarById(externalCalendarId);
  if (!eventCalendar || !externalCalendar) {
    throw new Error('同期先カレンダーを取得できません。システム管理の「年度更新設定」でC15/C16を確認してください。');
  }

  const holidayCalendars = CalendarApp.getCalendarsByName('日本の祝日');
  const holidayCalendar = holidayCalendars && holidayCalendars.length > 0 ? holidayCalendars[0] : null;
  if (!holidayCalendar) {
    Logger.log('[WARNING] 「日本の祝日」カレンダーが見つかりません。祝日スキップなしで同期します。');
  }

  // カラム情報を定義
  const eventColumns = [{ titleCol: 4 }]; // D列
  const externalColumns = [{ titleCol: 13 }]; // M列

  // 行事予定と対外行事のイベント同期処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[1];

    if (!date) continue; // 日付が空の場合はスキップ

    processEventUpdates(sheet, eventCalendar, eventColumns, row, date, "行事予定", i + 1, holidayCalendar);
    processEventUpdates(sheet, externalCalendar, externalColumns, row, date, "対外行事", i + 1, holidayCalendar);
  }
  Logger.log("[INFO] カレンダーの同期が完了しました。");
}

function processEventUpdates(sheet, calendar, columns, row, date, eventType, rowIndex, holidayCalendar) {
  if (!calendar) {
    Logger.log(`[WARNING] ${eventType}カレンダーが取得できないため、行 ${rowIndex} をスキップしました。`);
    return;
  }

  const holidays = holidayCalendar ? holidayCalendar.getEventsForDay(date) : [];
  const existingEvents = calendar.getEventsForDay(date) || [];
  const newEvents = [];
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
  const existingEventMap = {};
  const managedExistingEventMap = {};
  existingEvents.forEach(event => {
    const key = buildCalendarEventKey(event.getTitle(), event.getStartTime(), event.getEndTime());
    existingEventMap[key] = event;
    if (isManagedCalendarEvent(event)) {
      managedExistingEventMap[key] = event;
    }
  });

  // 新しいイベントをマッピング
  const newEventMap = {};
  newEvents.forEach(eventObj => {
    const key = buildCalendarEventKey(eventObj.title, eventObj.startTime, eventObj.endTime);
    newEventMap[key] = eventObj;
  });

  // 既存イベントの削除（このスクリプトが管理しているイベントのみ）
  for (const key in managedExistingEventMap) {
    if (!newEventMap[key]) {
      managedExistingEventMap[key].deleteEvent();
      eventsChanged = true;
      Logger.log(`[INFO] 削除された${eventType}イベント: タイトル "${managedExistingEventMap[key].getTitle()}"`);
    }
  }

  // 新しいイベントの作成
  for (const key in newEventMap) {
    if (!existingEventMap[key]) {
      const eventObj = newEventMap[key];
      let createdEvent;
      if (eventObj.isAllDay) {
        createdEvent = calendar.createAllDayEvent(eventObj.title, eventObj.startTime, eventObj.endTime);
      } else {
        createdEvent = calendar.createEvent(eventObj.title, eventObj.startTime, eventObj.endTime);
      }

      if (createdEvent) {
        markCalendarEventAsManaged(createdEvent);
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

function buildCalendarEventKey(title, startTime, endTime) {
  return title + '_' + startTime.getTime() + '_' + endTime.getTime();
}

function isManagedCalendarEvent(event) {
  const description = event.getDescription() || '';
  return description.indexOf(CALENDAR_SYNC_MANAGED_MARKER) !== -1;
}

function markCalendarEventAsManaged(event) {
  const description = event.getDescription() || '';
  if (description.indexOf(CALENDAR_SYNC_MANAGED_MARKER) !== -1) {
    return;
  }
  event.setDescription(description ? description + '\n' + CALENDAR_SYNC_MANAGED_MARKER : CALENDAR_SYNC_MANAGED_MARKER);
}

// convertFullWidthToHalfWidth() は common.js で定義されているため削除
