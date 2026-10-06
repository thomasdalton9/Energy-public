"""One-off: daily NYMEX Henry Hub front-month futures (EIA natural-gas/pri/fut, series RNGC1, USD/MMBtu)
from 2024-12-01, saved to discovery_archive/results/henry_hub/hh_front_month_daily.csv for the Gulf daily charts."""
import os, sys
import pandas as pd, requests

KEY = os.environ["EIA_API_KEY"]
rows, off = [], 0
while True:
    r = requests.get("https://api.eia.gov/v2/natural-gas/pri/fut/data", timeout=60, params={
        "api_key": KEY, "frequency": "daily", "data[0]": "value", "facets[series][]": "RNGC1",
        "start": "2024-12-01", "sort[0][column]": "period", "sort[0][direction]": "asc",
        "length": 5000, "offset": off})
    r.raise_for_status()
    d = r.json()["response"]["data"]
    rows += d
    if len(d) < 5000:
        break
    off += 5000
if not rows:
    r = requests.get("https://api.eia.gov/v2/natural-gas/pri/fut/data", timeout=60, params={
        "api_key": KEY, "frequency": "daily", "data[0]": "value", "facets[series][]": "RNGC1",
        "sort[0][column]": "period", "sort[0][direction]": "desc", "length": 5})
    print("no rows since 2024-12-01; latest available:", r.status_code, r.text[:800])
    sys.exit(1)
df = pd.DataFrame(rows)[["period", "value"]].rename(columns={"period": "date", "value": "HH_front_month_USD_per_MMBtu"})
df["HH_front_month_USD_per_MMBtu"] = pd.to_numeric(df["HH_front_month_USD_per_MMBtu"])
out = "discovery_archive/results/henry_hub/hh_front_month_daily.csv"
os.makedirs(os.path.dirname(out), exist_ok=True)
df.to_csv(out, index=False)
print(len(df), df.date.min(), df.date.max()); print(df.head(3)); print(df.tail(3))
