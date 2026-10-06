"""One-off: daily NYMEX Henry Hub front-month futures close (Yahoo Finance chart API, NG=F continuous, USD/MMBtu)
from 2024-12-01 -> discovery_archive/results/henry_hub/hh_front_month_daily.csv for the Gulf daily charts.
(EIA's natural-gas/pri/fut RNGC1 stops on 2024-04-05, so it cannot cover 2025-26.)
NG=F is Yahoo's continuous front-month: it jumps at each contract roll (unadjusted)."""
import os, sys, time
import pandas as pd
from curl_cffi import requests as cr

p1 = int(pd.Timestamp("2024-12-01").timestamp()); p2 = int(time.time()) + 86400
s = cr.Session()
for host in ("query1", "query2"):
    r = s.get(f"https://{host}.finance.yahoo.com/v8/finance/chart/NG=F?period1={p1}&period2={p2}&interval=1d",
              impersonate="chrome", timeout=40, headers={"Accept": "application/json"})
    print(host, r.status_code)
    if r.status_code == 200:
        break
r.raise_for_status()
res = r.json()["chart"]["result"][0]
df = pd.DataFrame({"date": pd.to_datetime(res["timestamp"], unit="s").normalize(),
                   "HH_front_month_USD_per_MMBtu": res["indicators"]["quote"][0]["close"]}).dropna()
out = "discovery_archive/results/henry_hub/hh_front_month_daily.csv"
os.makedirs(os.path.dirname(out), exist_ok=True)
df.to_csv(out, index=False)
print(len(df), df.date.min(), df.date.max()); print(df.head(3)); print(df.tail(3))
