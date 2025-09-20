import datetime
import hashlib
import os
import sys
from zoneinfo import ZoneInfo

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


def send_calendar(mode, past, future):
    # 当日 00:00 (JST) を基準に、過去 past 日と未来 future 日の予定を取得
    today = datetime.datetime.now(tz=datetime.timezone(datetime.timedelta(hours=9))).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    start = today - datetime.timedelta(days=past)
    end = today + datetime.timedelta(days=future)

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    items = ns.GetDefaultFolder(9).Items  # 9=olFolderCalendar
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    restriction = (
        f"([Start] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M %p')}') "
        f"OR ([End] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %H:%M %p')}')"
    )
    restricted_items = items.Restrict(restriction)
    print(f"Found {len([0 for _ in restricted_items])} event(s) from {start} to {end}")

    print(f"Mode: {mode}")
    # --- JSON 作成 ---
    if mode == "busystatus" or mode == "workinghours":
        events_allday = []
        events_timed = []
        for item in restricted_items:
            tz = datetime.timezone(item.Start - item.StartUTC)

            subject = item.Subject
            start_d = item.Start.replace(tzinfo=tz).strftime("%Y-%m-%d")
            start_dt = item.Start.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z")
            end_d = item.End.replace(tzinfo=tz).strftime("%Y-%m-%d")
            end_dt = item.End.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z")
            isAllday = item.AllDayEvent
            uid = hashlib.sha256(str(mode + subject + start_dt + end_dt + str(isAllday)).encode()).hexdigest()
            busy_status = int(item.BusyStatus)

            if item.AllDayEvent:
                events_allday.append(
                    {
                        "summary": subject,
                        "start": start_d,
                        "end": end_d,
                        "isAllDay": isAllday,
                        "id": uid,
                    }
                )
            else:
                events_timed.append(
                    {
                        "summary": subject,
                        "start": start_dt,
                        "end": end_dt,
                        "isAllDay": isAllday,
                        "id": uid,
                        "status": int(item.BusyStatus),
                    }
                )
        if mode == "busystatus":
            events = events_allday + merge_events(events_timed)
        elif mode == "workinghours":
            events = events_allday + get_working_hours(events_timed)

    elif mode == "event":
        events = []
        for item in restricted_items:

            tz = datetime.timezone(item.Start - item.StartUTC)
            subject = item.Subject
            start_d = item.Start.replace(tzinfo=tz).strftime("%Y-%m-%d")
            start_dt = item.Start.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z")
            end_d = item.End.replace(tzinfo=tz).strftime("%Y-%m-%d")
            end_dt = item.End.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z")
            isAllday = item.AllDayEvent
            uid = hashlib.sha256(str(mode + subject + start_dt + end_dt + str(isAllday)).encode()).hexdigest()
            busy_status = int(item.BusyStatus)

            events.append(
                {
                    "summary": busy_status_label[busy_status],
                    "start": (start_d if isAllday else start_dt),
                    "end": (end_d if isAllday else end_dt),
                    "isAllDay": isAllday,
                    "id": uid,
                }
            )

    print(f"Sending {len(events)} event(s) to GAS")

    # ====== POST リクエスト送信 ======
    response = requests.post(
        GAS_URL,
        params={"token": SECRET_TOKEN},  # URLパラメータでトークン送信
        json=events,  # 本文に JSON を送信
    )
    print(response.text)


def merge_events(events):
    """時間帯が重なっている予定と連続している予定をマージする"""
    if not events:
        return []

    timeranges = {
        "free": P.empty(),  # 0 : "空き時間",
        "tentative": P.empty(),  # 1 : "仮の予定",
        "busy": P.empty(),  # 2 : "予定あり",
        "elsewhere": P.empty(),  # 4 : "他の場所",
    }

    # 重なっている時間帯のうち、優先度の高い予定を抽出
    occupied = P.empty()
    merged_events = []
    for event in events:
        start = datetime.datetime.fromisoformat(event["start"])
        end = datetime.datetime.fromisoformat(event["end"])
        status = event["status"]

        if status == 3:
            # 最優先: 不在
            merged_events.append(event)
            occupied |= P.closed(start, end)
        else:
            timeranges[busy_status_key[status]] |= P.closed(start, end)

    # 次優先: 予定あり、仮の予定、他の場所、空き時間
    for status in [2, 1, 4, 0]:
        timerange = timeranges[busy_status_key[status]] - occupied

        for rng in timerange:
            start = rng.lower.strftime("%Y-%m-%dT%H:%M:%S%z")
            end = rng.upper.strftime("%Y-%m-%dT%H:%M:%S%z")
            uid = hashlib.sha256((str(busy_status_key[status]) + start + end).encode()).hexdigest()

            merged_events.append(
                {
                    "summary": busy_status_label[status],
                    "start": start,
                    "end": end,
                    "isAllDay": False,
                    "id": uid,
                }
            )
        occupied |= timerange

    return merged_events


def get_working_hours(events):
    """勤務予定の開始・終了時刻を1日毎にマージする"""
    if not events:
        return []

    timerange_event = P.empty()  # 予定
    timerange_working = P.empty()  # 勤務時間

    # 重なっている時間帯のうち、優先度の高い予定を抽出
    occupied = P.empty()
    merged_events = []
    for event in events:
        status = event["status"]

        if status == 3:
            # 最優先: 不在
            merged_events.append(event)
            occupied |= P.closed(
                datetime.datetime.fromisoformat(event["start"]), datetime.datetime.fromisoformat(event["end"])
            )
        elif status in [1, 2, 4]:
            # 次優先: 予定あり、仮の予定、他の場所
            timerange_event |= P.closed(
                datetime.datetime.fromisoformat(event["start"]), datetime.datetime.fromisoformat(event["end"])
            )
        else:
            # 無視: 空き時間
            pass

    for i in range(len(timerange_event) - 1):
        # 間を埋める
        if timerange_event[i].upper.date() == timerange_event[i + 1].lower.date():
            timerange_working |= P.closed(timerange_event[i].lower, timerange_event[i + 1].upper)

    timerange_working -= occupied

    for rng in timerange_working:
        start = rng.lower.strftime("%Y-%m-%dT%H:%M:%S%z")
        end = rng.upper.strftime("%Y-%m-%dT%H:%M:%S%z")
        uid = hashlib.sha256(("working" + start + end).encode()).hexdigest()

        merged_events.append(
            {
                "summary": "勤務",
                "start": start,
                "end": end,
                "isAllDay": False,
                "id": uid,
            }
        )

    return merged_events


if __name__ == "__main__":
    args = sys.argv
    if len(args) > 3:
        send_calendar(args[1], int(args[2]), int(args[3]))
    else:
        send_calendar("workinghours", 7, 28)
