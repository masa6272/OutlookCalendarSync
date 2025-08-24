import datetime
import os

import portion as P
import requests
import win32com.client
from dotenv import load_dotenv

# --- .env 読み込み ---
load_dotenv()
GAS_URL = os.getenv("GAS_URL")
SECRET_TOKEN = os.getenv("SECRET_TOKEN")
ICS_FILE = os.getenv("ICS_FILE")

busy_status_label = {
    0: "空き時間",  # olFree
    1: "仮の予定",  # olTentative
    2: "予定あり",  # olBusy
    3: "不在",  # olOutOfOffice
    4: "他の場所",  # olWorkingElsewhere
}

busy_status_key = {
    0: "free",  # olFree
    1: "tentative",  # olTentative
    2: "busy",  # olBusy
    3: "ooo",  # olOutOfOffice
    4: "elsewhere",  # olWorkingElsewhere
}


def send_calendar():

    # 30日前から180日後までの予定を取得
    start = datetime.datetime.now() - datetime.timedelta(days=30)
    end = start + datetime.timedelta(days=7) + datetime.timedelta(days=180)

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    items = ns.GetDefaultFolder(9).Items  # 9=olFolderCalendar
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    restriction = (
        f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %H:%M %p')}'"
    )
    restricted_items = items.Restrict(restriction)

    # --- JSON 作成 ---
    events_allday = []
    events_timed = []
    for item in restricted_items:
        tz = datetime.timezone(item.Start - item.StartUTC)
        if item.AllDayEvent:
            events_allday.append(
                {
                    "summary": item.Subject,
                    "start": item.Start.replace(tzinfo=tz).strftime("%Y-%m-%d"),
                    "end": item.End.replace(tzinfo=tz).strftime("%Y-%m-%d"),
                    "isAllDay": item.AllDayEvent,
                }
            )
        else:
            events_timed.append(
                {
                    "summary": item.Subject,
                    "start": item.Start.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z"),
                    "end": item.End.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z"),
                    # "isAllDay": item.AllDayEvent,
                    "status": int(item.BusyStatus),
                }
            )
    events = events_allday + merge_events(events_timed)
    print(events)

    # ====== POST リクエスト送信 ======
    response = requests.post(
        GAS_URL,
        params={"token": SECRET_TOKEN},  # URLパラメータでトークン送信
        json=events,  # 本文に JSON を送信
    )
    print(response.status_code, response.text)


def merge_events(events):
    """時間帯が重なっている予定と連続している予定をマージする"""
    if not events:
        return []

    print(f"original: {events}")

    timeranges = {
        "free": P.empty(),  # 0 : "空き時間",
        "tentative": P.empty(),  # 1 : "仮の予定",
        "busy": P.empty(),  # 2 : "予定あり",
        "elsewhere": P.empty(),  # 4 : "他の場所",
    }
    timeranges_ooo = []  # 3 : "不在"

    for event in events:
        start = datetime.datetime.fromisoformat(event["start"])
        end = datetime.datetime.fromisoformat(event["end"])
        status = event["status"]

        if status != 3:
            timeranges[busy_status_key[status]] |= P.closed(start, end)
        else:
            timeranges_ooo.append({"summary": event["summary"], "timerange": P.closed(start, end)})

    print(f"timeranges: {timeranges}")
    print(f"timeranges_ooo: {timeranges_ooo}")

    # 重なっている時間帯のうち、優先度の高い予定を抽出
    occupied = P.empty()
    merged_events = []

    # 最優先: 不在
    for event in timeranges_ooo:
        print(f"ooo event: {event}")
        merged_events.append(
            {
                "summary": f"不在: {event['summary']}",
                "start": event["timerange"].lower.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "end": event["timerange"].upper.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "isAllDay": False,
            }
        )
        occupied |= event["timerange"]

    # 次優先: 予定あり、仮の予定、他の場所、空き時間
    for status in [2, 1, 4, 0]:
        timerange = timeranges[busy_status_key[status]] - occupied

        for rng in timerange:
            merged_events.append(
                {
                    "summary": busy_status_label[status],
                    "start": rng.lower.strftime("%Y-%m-%dT%H:%M:%S%z"),
                    "end": rng.upper.strftime("%Y-%m-%dT%H:%M:%S%z"),
                    "isAllDay": False,
                }
            )
        occupied |= timerange

    return merged_events


if __name__ == "__main__":
    send_calendar()
