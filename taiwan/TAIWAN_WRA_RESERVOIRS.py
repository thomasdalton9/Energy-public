"""
Taiwan reservoir water levels and storage from the Water Resources Agency (WRA, Ministry of Economic Affairs) open-data
platform -> output/Data and Chart Outputs/taiwan_reservoirs_daily.xlsx.

Source: https://opendata.wra.gov.tw/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679?format=JSON&limit=1000 (data.gov.tw dataset 45501
'Reservoir water-regime data', keyless). The API returns only the last ~24 hours of hourly or daily readings of about 60
reservoirs - no history: WRA's own history (fhy.wra.gov.tw) needs an API key. The run therefore keeps the last reading of
the previous Taiwan calendar day for each reservoir and the committed workbook is the history store; history starts with
the first run. Days the workflow does not run stay gaps (never filled).

Sheets: 'Daily storage' (date x reservoir id, effective storage as published, nominally 10,000 m3), 'Daily level' (m),
'Reservoirs' (id, name if known, first date). The WRA reservoir name table ('Reservoir daily operation' dataset) could not be
reached from GitHub Actions (data.wra.gov.tw does not resolve; fhy.wra.gov.tw/Api needs a key), so names are blank and
the charts are labelled by the WRA reservoir identifier. Fill the 'name' column of the 'Names' sheet in
taiwan/taiwan_reservoir_names.csv (id,name) from a sourced list to label them.

    python3 taiwan/TAIWAN_WRA_RESERVOIRS.py --out "output/Data and Chart Outputs/taiwan_reservoirs_daily.xlsx"
"""
import argparse
import os
import sys
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import xlsx_notes  # noqa: E402

URL = "https://opendata.wra.gov.tw/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
DEFAULT_OUT = os.path.join(os.path.dirname(HERE), "output", "Data and Chart Outputs", "taiwan_reservoirs_daily.xlsx")
NAMES_CSV = os.path.join(HERE, "taiwan_reservoir_names.csv")
TW = timezone(timedelta(hours=8))


def fetch():
    last = None
    for i in range(3):
        try:
            r = requests.get(URL, params={"format": "JSON", "limit": 1000}, headers=UA, timeout=(15, 90))
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            last = e
            print(f"  attempt {i + 1}/3: {type(e).__name__}", file=sys.stderr)
    raise last


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return float("nan")


def read_old(path, sheet):
    if not os.path.exists(path):
        return None
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index)
        d.columns = [int(c) for c in d.columns]
        return d
    except Exception as e:  # noqa: BLE001
        print(f"  stored sheet {sheet} unreadable ({type(e).__name__})", file=sys.stderr)
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    recs = fetch()
    df = pd.DataFrame(recs)
    df["t"] = pd.to_datetime(df["observationtime"], errors="coerce")
    df = df.dropna(subset=["t"])
    df["id"] = df["reservoiridentifier"].astype(int)
    df["storage"] = df["effectivewaterstoragecapacity"].map(num)
    df["level"] = df["waterlevel"].map(num)
    today = datetime.now(TW).date()
    yday = pd.Timestamp(today - timedelta(days=1))
    day = df[df["t"].dt.normalize() == yday].sort_values("t")
    print(f"  {len(df)} readings {df['t'].min()}..{df['t'].max()}; {day['id'].nunique()} reservoirs with readings on {yday:%Y-%m-%d}")
    new_s = day.groupby("id")["storage"].last().to_frame(yday).T
    new_l = day.groupby("id")["level"].last().to_frame(yday).T
    sheets = {}
    for name, new in (("Daily storage", new_s), ("Daily level", new_l)):
        old = read_old(args.out, name)
        if old is not None and not old.empty:
            new = pd.concat([old[~old.index.isin(new.index)], new]).sort_index()
        new.index.name = "date"
        sheets[name] = new.reindex(sorted(new.columns), axis=1)
    if sheets["Daily storage"].dropna(how="all").empty:
        raise SystemExit("no readings stored and none for yesterday")
    names = {}
    if os.path.exists(NAMES_CSV):
        n = pd.read_csv(NAMES_CSV)
        names = dict(zip(n["id"].astype(int), n["name"]))
    ids = sheets["Daily storage"].columns
    first = sheets["Daily storage"].apply(lambda c: c.first_valid_index())
    sheets["Reservoirs"] = pd.DataFrame({"name": [names.get(i, "") for i in ids], "first_date": first.values}, index=pd.Index(ids, name="id"))
    sheets["Reservoirs"]["name"] = [n if n else f"WRA {i}" for i, n in zip(ids, sheets["Reservoirs"]["name"])]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES, {"UNITS", "SOURCE", "UPDATES"})
    print(f"Saved {sheets['Daily storage'].shape} -> {args.out}")


NOTES = [
    "UNITS",
    "Daily storage: effective storage as published by WRA (field effectivewaterstoragecapacity; WRA convention is 10,000 m3, the API "
    "does not state the unit). Daily level: water level, m. Value = the last reading of the Taiwan calendar day (UTC+8).",
    "",
    "SOURCE",
    "Water Resources Agency, Ministry of Economic Affairs: reservoir water-regime open data "
    "(https://opendata.wra.gov.tw/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679 ; data.gov.tw dataset 45501). The API holds only the "
    "last ~24 hours, so the history starts with the first run. Reservoir names are not in this dataset; WRA's name table could not be reached.",
    "",
    "UPDATES",
    "Daily (the source keeps no history). Each run stores yesterday's last reading per reservoir; days not run are gaps.",
]

if __name__ == "__main__":
    main()
