"""Probe 4: gassco.eu realTimeGraphData?year&month&day - how far back it goes, how the timestamp relates to the requested day,
and a dump of daily values 2024-01-01..today for comparison with GB NTS. Prints only (CSV lines start with 'CSV,')."""
import datetime as dt
import time
from concurrent.futures import ThreadPoolExecutor

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URL = "https://gassco.eu/wp-json/gassco/v1/realTimeGraphData"
S = requests.Session()
S.headers.update(H)


def get(d):
    for k in range(4):
        try:
            r = S.get(URL, params={"year": d.year, "month": d.month, "day": d.day}, timeout=40)
            if r.status_code == 200:
                if not r.text.strip():
                    return None
                j = r.json()
                v = {n["name"]: n["value"] for n in j["nodes"]}
                return v, j["timestamp"]
        except Exception as e:  # noqa: BLE001
            err = type(e).__name__
        time.sleep(1 + k)
    return "ERR"


def line(d, res):
    if res is None:
        return f"{d},EMPTY"
    if res == "ERR":
        return f"{d},ERR"
    v, ts = res
    return f"{d},{v.get('Germany')},{v.get('France')},{v.get('Belgium')},{v.get('Great Britain')},{v.get('Other')},{v.get('Total')},{dt.datetime.utcfromtimestamp(ts/1000):%Y-%m-%d %H:%M}"


print("--- A: 15th of each month 2008..2022")
ds = [dt.date(y, m, 15) for y in range(2008, 2023) for m in range(1, 13)]
with ThreadPoolExecutor(6) as ex:
    for d, res in zip(ds, ex.map(get, ds)):
        print("A", line(d, res))

print("--- B: edge days")
today = dt.datetime.utcnow().date()
for d in [today - dt.timedelta(days=k) for k in range(-1, 6)] + [dt.date(2026, 9, k) for k in (1, 2, 30)] + [dt.date(2026, 2, 28), dt.date(2026, 3, 1)]:
    print("B", line(d, get(d)))

print("--- C: daily dump")
ds = [dt.date(2024, 1, 1) + dt.timedelta(days=i) for i in range((today - dt.date(2024, 1, 1)).days + 1)]
with ThreadPoolExecutor(6) as ex:
    for d, res in zip(ds, ex.map(get, ds)):
        print("CSV," + line(d, res))
