import datetime
import os

import requests
import win32com.client
from dotenv import load_dotenv

# --- .env 読み込み ---
load_dotenv()
GAS_URL = os.getenv("GAS_URL")
SECRET_TOKEN = os.getenv("SECRET_TOKEN")
ICS_FILE = os.getenv("ICS_FILE")

busy_status = {
    0: "空き時間",  # olFree
    1: "仮の予定",  # olTentative
    2: "予定あり",  # olBusy
    3: "不在",  # olOutOfOffice
    4: "他の場所",  # olWorkingElsewhere
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
                    # "summary": item.Subject,
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
    # 区切りとなる時刻を抽出
    timestamps_boundary = []
    for event in events:
        start = datetime.datetime.fromisoformat(event["start"])
        end = datetime.datetime.fromisoformat(event["end"])
        status = event["status"]
        timestamps_boundary.append({"timestamp": start, "status": status})
        timestamps_boundary.append({"timestamp": end, "status": -1})
    # 時刻でソート
    timestamps_boundary.sort(key=lambda x: x["timestamp"])

    print(f"boundary: {timestamps_boundary}")

    # boundaryで区切って、全部の時間帯を抽出
    timerange_status_all = []
    for event in events:
        start = datetime.datetime.fromisoformat(event["start"])
        end = datetime.datetime.fromisoformat(event["end"])
        status = event["status"]

        for boundary in timestamps_boundary:
            boundary_ts = boundary["timestamp"]
            if start < boundary_ts < end:
                # start - boundary, boundary - end に分割
                timerange_status_all.append({"start": start, "end": boundary_ts, "status": status})
                start = boundary_ts

        timerange_status_all.append({"start": start, "end": end, "status": status})
    timerange_status_all.sort(key=lambda x: x["start"])
    print(f"all: {timerange_status_all}")

    # 重なっている時間帯のうち、優先度の高い予定を抽出
    timerange_status = []

    prev_start = timerange_status_all[0]["start"]
    prev_end = timerange_status_all[0]["end"]
    prev_status = timerange_status_all[0]["status"]

    for i in range(len(timerange_status_all) - 1):
        next_start = timerange_status_all[i + 1]["start"]
        next_end = timerange_status_all[i + 1]["end"]
        next_status = timerange_status_all[i + 1]["status"]

        if prev_end <= next_start:
            # 重なっていない
            timerange_status.append({"start": prev_start, "end": prev_end, "status": prev_status})

            if i == len(timerange_status_all) - 2:
                # 最後のnextを追加して終了
                timerange_status.append({"start": next_start, "end": next_end, "status": next_status})
                break
            else:
                prev_start = next_start
                prev_end = next_end
                prev_status = next_status

        else:
            # 重なっている場合、必ず時間帯が同じ
            assert prev_start == next_start
            assert prev_end == next_end
            status_prior = compare_status(prev_status, next_status)

            if i == len(timerange_status_all) - 2:
                # ラスト
                timerange_status.append({"start": prev_start, "end": prev_end, "status": status_prior})
            else:
                prev_status = status_prior

    timerange_status.sort(key=lambda x: x["start"])

    print(f"timerange: {timerange_status}")

    # 連続する時間帯をマージ
    merged_status = []

    prev_start = timerange_status[0]["start"]
    prev_end = timerange_status[0]["end"]
    prev_status = timerange_status[0]["status"]

    for i in range(len(timerange_status) - 1):
        next_start = timerange_status[i + 1]["start"]
        next_end = timerange_status[i + 1]["end"]
        next_status = timerange_status[i + 1]["status"]

        if prev_end < next_start or prev_status != next_status:
            # prevは確定
            if prev_status != -1:  # 予定なしは追加しない
                merged_status.append(
                    {
                        "start": prev_start.strftime("%Y-%m-%dT%H:%M:%S%z"),
                        "end": prev_end.strftime("%Y-%m-%dT%H:%M:%S%z"),
                        "status": prev_status,
                    }
                )

            if i < len(timerange_status) - 2:
                # 次 (next) に進む
                prev_start = next_start
                prev_end = next_end
                prev_status = next_status
            else:
                # 最後のnextを追加して終了
                if next_status != -1:
                    merged_status.append(
                        {
                            "start": next_start.strftime("%Y-%m-%dT%H:%M:%S%z"),
                            "end": next_end.strftime("%Y-%m-%dT%H:%M:%S%z"),
                            "status": next_status,
                        }
                    )
                break

        else:
            # prevとnextをマージ
            # start: prev_startのまま
            # end: next_endに更新
            prev_end = next_end
            # status: prev_statusのまま

    print(f"status: {merged_status}")

    # Eventの形式に変換
    merged_events = [
        {
            "summary": busy_status[event["status"]],
            "start": event["start"],
            "end": event["end"],
            "isAllDay": False,
        }
        for event in merged_status
    ]

    return merged_events


def compare_status(status1, status2):
    """2つのBusyStatusを比較し、優先度の高い方を返す"""
    # ↑ 最低優先度
    # -1: "予定なし"
    # 0 : "空き時間",
    # 4 : "他の場所",
    # 1 : "仮の予定",
    # 2 : "予定あり",
    # 3 : "不在",
    # ↓ 最高優先度

    if status1 == 3 or status2 == 3:
        return 3
    elif status1 == 2 or status2 == 2:
        return 2
    elif status1 == 1 or status2 == 1:
        return 1
    elif status1 == 4 or status2 == 4:
        return 4
    elif status1 == 0 or status2 == 0:
        return 0
    else:
        return -1


if __name__ == "__main__":
    send_calendar()
