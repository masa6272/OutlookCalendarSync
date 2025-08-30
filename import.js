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

    log = "Start\n";

    // --- Google カレンダー取得 ---
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);

    // 既存イベントをIDでマッピング
    const existingEvents = cal.getEvents(
      new Date("2000-01-01"),
      new Date("2100-01-01")
    );

    const existingMap = {};
    existingEvents.forEach((ev) => {
      const id = ev.getTag("outlookId") || ""; // カスタムタグでIDを保持、なければ空文字列
      if (id) existingMap[id] = ev;
      else ev.deleteEvent(); // タグがないイベントは削除
    });
    log += `Existing: ${Object.keys(existingMap).length} event(s)\n`;

    // JSON パース
    const incomingEvents = JSON.parse(e.postData.contents);
    // 受信イベントをIDでマッピング
    const incomingMap = {};
    incomingEvents.forEach((ev) => (incomingMap[ev.id] = ev));

    log += `Incoming: ${Object.keys(incomingMap).length} event(s)\n`;

    // 追加・更新
    updated = 0;
    added = 0;
    incomingEvents.forEach((ev) => {
      const existing = existingMap[ev.id];

      if (!existing) {
        // 新規作成
        let newEv;
        if (ev.isAllDay) {
          newEv = cal.createAllDayEvent(
            ev.summary,
            new Date(ev.start),
            new Date(ev.end)
          );
        } else {
          newEv = cal.createEvent(
            ev.summary,
            new Date(ev.start),
            new Date(ev.end)
          );
        }
        newEv.setTag("outlookId", ev.id);
        added++;
      } else {
        // 更新（タイトルや時間が違う場合のみ）
        if (
          existing.getTitle() !== ev.summary ||
          existing.getStartTime().getTime() !== new Date(ev.start).getTime() ||
          existing.getEndTime().getTime() !== new Date(ev.end).getTime()
        ) {
          existing.setTitle(ev.summary);
          existing.setTime(new Date(ev.start), new Date(ev.end));
          updated++;
        }
      }
    });
    log += `Added: ${added} event(s)\n`;
    log += `Updated: ${updated} event(s)\n`;

    // 削除（Googleカレンダーにあるが、受信データにないID）
    deleted = 0;
    Object.keys(existingMap).forEach((id) => {
      if (!incomingMap[id]) {
        existingMap[id].deleteEvent();
        deleted++;
      }
    });
    log += `Deleted: ${deleted} event(s)\n`;

    log += "End\n";
    return ContentService.createTextOutput(log).setMimeType(
      ContentService.MimeType.TEXT
    );
  } catch (err) {
    return ContentService.createTextOutput(
      "Error: " + err.message + "\n" + log
    ).setMimeType(ContentService.MimeType.TEXT);
  }
}
