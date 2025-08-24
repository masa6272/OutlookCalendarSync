// === Webhook 本体 ===
function doPost(e) {
  try {
    // --- スクリプトプロパティから秘密情報取得 ---
    const props = PropertiesService.getScriptProperties();
    const SECRET_TOKEN = props.getProperty("SECRET_TOKEN");
    const CALENDAR_ID = props.getProperty("CALENDAR_ID");

    // --- 認証チェック ---
    const token = e.parameter.token;
    if (token !== SECRET_TOKEN) {
      return ContentService.createTextOutput("Unauthorized").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // --- ICS データを取得 ---
    const icsData = e.postData.contents;
    if (!icsData)
      return ContentService.createTextOutput("No ICS data").setMimeType(
        ContentService.MimeType.TEXT
      );

    // --- Google カレンダー取得 ---
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);

    // --- 既存イベントを全削除 ---
    const allEvents = cal.getEvents(
      new Date("2000-01-01"),
      new Date("2100-01-01")
    );
    allEvents.forEach((ev) => ev.deleteEvent());

    // --- ICS パース & Googleカレンダーに登録 ---
    const events = icsData.split("BEGIN:VEVENT").slice(1);
    events.forEach((eventText) => {
      const dtStartMatch = eventText.match(
        /DTSTART(?:;TZID=[^:]*)?:(\d{8}T\d{6}(?:[+-]\d{4})?)/
      );
      const dtEndMatch = eventText.match(
        /DTEND(?:;TZID=[^:]*)?:(\d{8}T\d{6}(?:[+-]\d{4})?)/
      );
      const summaryMatch = eventText.match(/SUMMARY:(.*)/);

      if (dtStartMatch && dtEndMatch && summaryMatch) {
        const start = parseICSTime(dtStartMatch[1]);
        const end = parseICSTime(dtEndMatch[1]);
        const title = summaryMatch[1];

        if (start && end) cal.createEvent(title, start, end);
      }
    });

    return ContentService.createTextOutput("OK").setMimeType(
      ContentService.MimeType.TEXT
    );
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message).setMimeType(
      ContentService.MimeType.TEXT
    );
  }
}

// === ICS 日時文字列を JS Date に変換 ===
function parseICSTime(icsTime) {
  // YYYYMMDDTHHMMSSZ または ±hhmm 付き
  const match = icsTime.match(
    /(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z|[+-]\d{4})?/
  );
  if (!match) return null;

  const [, year, month, day, hour, min, sec, tz] = match;

  // UTC Date として作成
  let date = new Date(
    Date.UTC(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day),
      parseInt(hour),
      parseInt(min),
      parseInt(sec)
    )
  );

  // タイムゾーン補正
  if (tz && tz !== "Z") {
    const sign = tz[0] === "+" ? 1 : -1;
    const offsetHours = parseInt(tz.slice(1, 3));
    const offsetMinutes = parseInt(tz.slice(3, 5));
    const offsetMs = sign * ((offsetHours * 60 + offsetMinutes) * 60 * 1000);
    date = new Date(date.getTime() - offsetMs); // JS Date はローカル時間に変換
  }

  return date;
}
