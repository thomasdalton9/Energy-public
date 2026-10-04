"""
Estonia's daily gas flows from Elering (the Estonian TSO) Live Dashboard API (free, no key):

  output/Data and Chart Outputs/elering_gas_daily.xlsx
    sheet "Daily": date (gas day, 06:00-06:00 CET), GWh per day. Sign: positive = gas entering Estonia, negative = leaving.
        EE_balticconnector   Balticconnector (Finland <-> Estonia): + from Finland, - to Finland
        EE_karksi            Karksi (Latvia <-> Estonia): + from Latvia, - to Latvia
        EE_narva, EE_varska  Russia border points (idle since 2022)
        EE_consumption       gas delivered from the transmission network to Estonian consumers (positive number)
    sheet "Units": source and definitions

Why: ENTSOG has no Karksi row (Latvia's Karksi exit is blank), so the Latvian and Estonian balances missed the Latvia -> Estonia -> Finland
flow of Inčukalns gas (Estonian consumption plus the Balticconnector exports to Finland).

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days plus any gap; history from 2021-10-01.

Usage: python3 ELERING_GAS_DAILY.py [--out-dir DIR] [--start 2021-10-01]
"""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "elering_gas_daily.xlsx"
BASE = "https://dashboard.elering.ee/api"
REVISION_DAYS = 14
POINTS = {"bc": "EE_balticconnector", "karksi": "EE_karksi", "narva": "EE_narva", "varska": "EE_varska"}
COLUMNS = list(POINTS.values()) + ["EE_consumption"]


def get(path, start, end):
    for i in range(4):
        try:
            r = requests.get(BASE + path, params={"start": start.strftime("%Y-%m-%dT00:00:00.000Z"), "end": end.strftime("%Y-%m-%dT00:00:00.000Z")}, timeout=90)
            if r.ok:
                return r.json().get("data")
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError("gave up on Elering")


def gas_day(ts):
    """Hourly UTC epoch seconds -> gas day (starts 06:00 local time of the CET zone, i.e. 05:00/04:00 UTC)."""
    t = pd.to_datetime(ts, unit="s", utc=True).dt.tz_convert("Europe/Berlin") - pd.Timedelta(hours=6)
    return t.dt.tz_localize(None).dt.normalize()


def pull_chunk(d0, d1):
    out = {}
    cb = get("/gas-transmission/cross-border", d0, d1) or {}
    for key, col in POINTS.items():
        rows = cb.get(key) or []
        if rows:
            df = pd.DataFrame(rows)
            df["day"] = gas_day(df["timestamp"])
            out[col] = pd.to_numeric(df["volume"], errors="coerce").groupby(df["day"]).sum(min_count=1) / 1e6      # kWh -> GWh
    gs = get("/gas-system", d0, d1) or []
    if gs:
        df = pd.DataFrame(gs)
        df["day"] = gas_day(df["timestamp"])
        out["EE_consumption"] = pd.to_numeric(df["value"], errors="coerce").groupby(df["day"]).sum(min_count=1) / 1e6
    return pd.DataFrame(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-10-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = pd.DataFrame(columns=COLUMNS)
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.dropna(subset=["date"]).set_index("date").sort_index()
        except Exception as e:  # noqa: BLE001
            print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
            old = pd.DataFrame(columns=COLUMNS)
    have = old.dropna(how="all")
    fs = start if have.empty or not set(COLUMNS) <= set(old.columns) else max(start, have.index.max().date() - timedelta(days=REVISION_DAYS))
    parts, d = [], fs
    while d <= today:
        e = min(d + timedelta(days=31), today + timedelta(days=1))
        p = pull_chunk(datetime.combine(d, datetime.min.time()), datetime.combine(e, datetime.min.time()))
        if len(p):
            parts.append(p)
        print(f"  {d}..{e}: {len(p)} days", flush=True)
        d = e
    if not parts:
        print("no rows")
        return
    new = pd.concat(parts)
    new = new.groupby(new.index).last().sort_index()
    new = new.reindex(columns=COLUMNS)
    new = new[new.index.date < today]            # drop the gas day in progress
    comb = new.combine_first(old[[c for c in COLUMNS if c in old]]) if len(old) else new
    comb.loc[new.index, COLUMNS] = new[COLUMNS]
    comb = comb.sort_index().round(3)
    comb.index.name = "date"
    print(f"workbook {comb.index.min():%Y-%m-%d} .. {comb.index.max():%Y-%m-%d}, {len(comb)} days")
    print("TWh per year:\n" + (comb.groupby(comb.index.year).sum() / 1000).round(2).T.to_string())
    lines = ["Estonia - daily gas flows (Elering)", "",
             "Source", "Elering AS (Estonian gas TSO), Live Dashboard API https://dashboard.elering.ee/api (gas-transmission/cross-border and gas-system). Free, no key.",
             "", "Units and definitions",
             "GWh per gas day (the source is hourly kWh, summed over the 06:00-06:00 CET gas day). Positive = gas entering Estonia, negative = leaving. "
             "EE_balticconnector = Finland (+) / to Finland (-); EE_karksi = from Latvia (+) / to Latvia (-); EE_narva and EE_varska = Russian "
             "border points (idle since 2022); EE_consumption = gas delivered from the transmission network to Estonian consumers (a positive number). "
             "ENTSOG lists no Karksi flow, so this fills the Latvia - Estonia link.",
             f"Re-fetches the last {REVISION_DAYS} days each run; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} days, "
             f"{comb.index.min():%Y-%m-%d} to {comb.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
