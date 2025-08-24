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

    // JSON パース
    const events = JSON.parse(e.postData.contents);

    // --- Google カレンダー取得 ---
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);

    // --- 既存イベントを全削除 ---
    const allEvents = cal.getEvents(
      new Date("2000-01-01"),
      new Date("2100-01-01")
    );
    allEvents.forEach((ev) => ev.deleteEvent());

    // 登録
    events.forEach((ev) => {
      if (ev.isAllDay) {
        cal.createAllDayEvent(ev.summary, new Date(ev.start), new Date(ev.end));
      } else {
        cal.createEvent(ev.summary, new Date(ev.start), new Date(ev.end), {});
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
