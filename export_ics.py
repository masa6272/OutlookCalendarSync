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
    events = []
    for item in restricted_items:
        tz = datetime.timezone(item.Start - item.StartUTC)
        if item.AllDayEvent:
            ev = {
                "summary": (item.Subject),
                "start": item.Start.replace(tzinfo=tz).strftime("%Y-%m-%d"),
                "end": item.End.replace(tzinfo=tz).strftime("%Y-%m-%d"),
                "isAllDay": item.AllDayEvent,
            }
        else:
            ev = {
                "summary": "予定あり",
                "start": item.Start.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z"),
                "end": item.End.replace(tzinfo=tz).strftime("%Y-%m-%dT%H:%M:%S%z"),
                "isAllDay": item.AllDayEvent,
            }
        events.append(ev)
    print(events)

    # ====== POST リクエスト送信 ======
    response = requests.post(
        GAS_URL,
        params={"token": SECRET_TOKEN},  # URLパラメータでトークン送信
        json=events,  # 本文に JSON を送信
    )
    print(response.status_code, response.text)


if __name__ == "__main__":
    send_calendar()
