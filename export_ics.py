import datetime
import os

import requests
import win32com.client
from dotenv import load_dotenv
from ics import Calendar, Event

# --- .env 読み込み ---
load_dotenv()
GAS_URL = os.getenv("GAS_URL")
SECRET_TOKEN = os.getenv("SECRET_TOKEN")
ICS_FILE = os.getenv("ICS_FILE")

# 30日前から180日後までの予定を取得
start = datetime.datetime.now() - datetime.timedelta(days=30)
end = start + datetime.timedelta(days=7) + datetime.timedelta(days=180)

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
items = ns.GetDefaultFolder(9).Items  # 9=olFolderCalendar
items.IncludeRecurrences = True
items.Sort("[Start]")

restriction = f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %H:%M %p')}'"
restricted_items = items.Restrict(restriction)

c = Calendar()

for item in restricted_items:
    print(item.Subject, item.Start, item.End, item.AllDayEvent)
    print(item.EntryID, item.AllDayEvent)
    print(item.StartUTC, item.EndUTC, item.StartTimeZone, item.EndTimeZone)
    tz = datetime.timezone(item.Start - item.StartUTC)
    print("tz:", tz)

    start_local = item.Start
    start_tz = item.Start.replace(tzinfo=tz)
    end_local = item.End
    end_tz = item.End.replace(tzinfo=tz)

    print("start:", start_local, start_tz)
    print("end:", end_local, end_tz)

    if item.AllDayEvent:
        summary = item.Subject
    else:
        summary = "予定あり"

    e = Event()
    e.name = summary
    e.begin = start_tz
    e.end = end_tz

    print("e.begin:", e.begin)
    print("e.end:", e.end)

    # Outlook EntryIDをUIDに入れて後で差分判定に利用
    e.uid = item.EntryID
    c.events.add(e)

with open(ICS_FILE, "w", encoding="utf-8") as f:
    f.writelines(c)

# ====== ICS ファイル読み込み ======
with open(ICS_FILE, "r", encoding="utf-8") as f:
    ics_data = f.read()

# ====== POST リクエスト送信 ======
response = requests.post(
    GAS_URL,
    params={"token": SECRET_TOKEN},  # URLパラメータでトークン送信
    data=ics_data.encode("utf-8"),  # 本文に ICS を送信
    headers={"Content-Type": "text/plain"},  # プレーンテキストとして送信
)

print(response.status_code)
print(response.text)
