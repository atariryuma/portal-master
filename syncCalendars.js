/**
 * @fileoverview Googleカレンダー同期機能
 * @description 年間行事予定表の内容をGoogleカレンダーに同期し、
 *              校内行事・対外行事を別々のカレンダーに登録します。
 *              時間指定、全日イベント、祝日の自動スキップに対応。
 */

const CALENDAR_SYNC_MANAGED_MARKER = '[PORTAL_MASTER_MANAGED]';

function syncCalendars() {
  const sheet = getAnnualScheduleSheetOrThrow();
  const data = sheet.getDataRange().getValues();

  const eventCalendarId = getOrCreateCalendarId('EVENT');
  const externalCalendarId = getOrCreateCalendarId('EXTERNAL');

  Logger.log("[INFO] カレンダーID取得完了、同期処理を開始します。");

  const eventCalendar = CalendarApp.getCalendarById(eventCalendarId);
  const externalCalendar = CalendarApp.getCalendarById(externalCalendarId);
  if (!eventCalendar || !externalCalendar) {
    throw new Error('同期先カレンダーを取得できません。システム管理の「年度更新設定」でC15/C16を確認してください。');
  }

  const holidayCalendars = CalendarApp.getCalendarsByName(ANNUAL_SCHEDULE.HOLIDAY_CALENDAR_NAME);
  const holidayCalendar = holidayCalendars && holidayCalendars.length > 0 ? holidayCalendars[0] : null;
  if (!holidayCalendar) {
    Logger.log('[WARNING] 「日本の祝日」カレンダーが見つかりません。祝日スキップなしで同期します。');
  }

  const eventColumns = [{ titleCol: ANNUAL_SCHEDULE.INTERNAL_EVENT_COLUMN }];
  const externalColumns = [{ titleCol: ANNUAL_SCHEDULE.EXTERNAL_EVENT_COLUMN }];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = normalizeToDate(row[ANNUAL_SCHEDULE.DATE_INDEX]);

    if (!date) continue;

    processEventUpdates(eventCalendar, eventColumns, row, date, "行事予定", i + 1, holidayCalendar);
    processEventUpdates(externalCalendar, externalColumns, row, date, "対外行事", i + 1, holidayCalendar);
  }
  Logger.log("[INFO] カレンダーの同期が完了しました。");
}

function processEventUpdates(calendar, columns, row, date, eventType, rowIndex, holidayCalendar) {
  try {
    if (!calendar) {
      Logger.log(`[WARNING] ${eventType}カレンダーが取得できないため、行 ${rowIndex} をスキップしました。`);
      return;
    }

    const holidays = holidayCalendar ? holidayCalendar.getEventsForDay(date) : [];
    const existingEvents = calendar.getEventsForDay(date) || [];
    const newEvents = [];
    let eventsChanged = false;

    columns.forEach(function(colDef) {
      // バッチ読み取り済みの data[row] から直接取得（個別 getRange().getValue() を排除）
      const titleCell = row[colDef.titleCol - 1];
      if (titleCell) {
        const titles = String(titleCell)
          .split('\n')
          .map(function(t) { return t.trim().replace(/^・/, ''); })
          .filter(function(t) { return t; });

        titles.forEach(function(title) {
          const isHoliday = holidays.some(function(holiday) {
            return holiday.getTitle() === title ||
              holiday.getTitle() === "振替休日";
          });
          if (!isHoliday) {
            const eventInfo = parseEventTimesAndDates(title, date);
            if (eventInfo) {
              newEvents.push(eventInfo);
            }
          }
        });
      }
    });

    const existingEventMap = {};
    const managedExistingEventMap = {};
    existingEvents.forEach(function(event) {
      const key = buildCalendarEventKey(event.getTitle(), event.getStartTime(), event.getEndTime());
      existingEventMap[key] = event;
      if (isManagedCalendarEvent(event)) {
        managedExistingEventMap[key] = event;
      }
    });

    const newEventMap = {};
    newEvents.forEach(function(eventObj) {
      const key = buildCalendarEventKey(eventObj.title, eventObj.startTime, eventObj.endTime);
      newEventMap[key] = eventObj;
    });

    Object.keys(managedExistingEventMap).forEach(function(key) {
      if (!newEventMap[key]) {
        managedExistingEventMap[key].deleteEvent();
        eventsChanged = true;
        Logger.log(`[INFO] 削除された${eventType}イベント: タイトル "${managedExistingEventMap[key].getTitle()}"`);
      }
    });

    Object.keys(newEventMap).forEach(function(key) {
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
    });

    if (eventsChanged) {
      Logger.log(`[INFO] ${eventType}イベントの変更が完了しました。日付: ${date}`);
    }
  } catch (error) {
    Logger.log(`[ERROR] ${eventType}イベント処理中にエラー（行 ${rowIndex}）: ${error.toString()}`);
  }
}

function parseEventTimesAndDates(title, date) {
  const trimmedTitle = convertFullWidthToHalfWidth(title.trim());
  const originalTitle = trimmedTitle;

  let isAllDay = false;
  let startTime = new Date(date);
  let endTime = null;

  const timePatternRange = /(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?(?:\s*[~～]\s*)(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?/;
  const timePatternSingle = /(\d{1,2})[:：時](\d{2}|\d{1,2}分?|半)?/;

  let timeMatch = trimmedTitle.match(timePatternRange);
  if (timeMatch) {
    const [_, startHour, startMinute, endHour, endMinute] = timeMatch;
    startTime = setEventTime(startTime, startHour, startMinute);
    endTime = setEventTime(new Date(date), endHour, endMinute);

    if (endTime <= startTime) {
      endTime.setDate(endTime.getDate() + 1);
    }

    isAllDay = false;
    return {startTime, endTime, title: originalTitle, isAllDay};
  }

  timeMatch = trimmedTitle.match(timePatternSingle);
  if (timeMatch) {
    const [_, startHour, startMinute] = timeMatch;
    startTime = setEventTime(startTime, startHour, startMinute);
    endTime = new Date(startTime.getTime() + 60 * 60 * 1000);

    isAllDay = false;
    return {startTime, endTime, title: originalTitle, isAllDay};
  }

  isAllDay = true;
  endTime = new Date(startTime.getTime() + 24 * 60 * 60 * 1000);
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
