"""
Daily physical flow at the ENTSOG entry points whose far side is Norway (the points where Norwegian pipeline gas is reported):

  output/Data and Chart Outputs/entsog_norway_entries_daily.xlsx
    sheet "Daily": date (gas day), GWh per day, one column per point: <CC>_<point label> (e.g. DE_Dornum GASPOOL, BE_Zeebrugge ZPT,
                   FR_Dunkerque, UK_Easington, UK_St. Fergus, DK_North Sea Entry)
    sheet "Units": source and definitions

Why: ENTSOG's Emden points (EPT1, EPT2, NPT) report no data under any indicator, so Germany's Norwegian imports in the main ENTSOG pull
are only Dornum. This workbook gives the ENTSOG-visible Norwegian volume per receiving country; the master adds Gassco's flow to Germany
minus the Dornum volume as the missing Emden component. UK Easington and St Fergus also carry UK North Sea gas.

Incremental: reads the committed workbook, re-fetches the last 45 days plus any gap; history from 2021-10-04 (ENTSOG keeps ~5 years).

Usage: python3 ENTSOG_NORWAY_ENTRIES_DAILY.py [--out-dir DIR] [--start 2021-10-04]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "entsog_norway_entries_daily.xlsx"
RELOAD_DAYS = 45


def get(path, params, tries=4):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(10 * (i + 1))
        except requests.RequestException:
            time.sleep(10 * (i + 1))
    raise RuntimeError(f"ENTSOG {path} failed")


def norway_entries():
    """(operatorKey, pointKey) -> column name, for every entry whose adjacent system is in Norway and whose operator is not Norwegian."""
    out = {}
    for i in get("interconnections", {"limit": -1}).get("interconnections", []):
        opk = i.get("toOperatorKey")
        if opk and i.get("fromCountryKey") == "NO" and not opk.startswith("NO"):
            out[(opk, i["pointKey"])] = f"{opk[:2]}_{i.get('pointLabel')}"
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-10-04")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = date.fromisoformat(args.start)
    end = date.today() - timedelta(days=1)
    old = pd.DataFrame()
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.dropna(subset=["date"]).set_index("date").sort_index()
        except Exception as e:  # noqa: BLE001
            print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
    fs = start if old.empty else max(start, (old.index.max() - timedelta(days=RELOAD_DAYS)).date())
    cols = norway_entries()
    print(f"{len(cols)} Norway-adjacent entry points; fetching {fs} -> {end}", flush=True)
    series = {}
    for (opk, pk), name in sorted(cols.items()):
        rec, cur = {}, fs
        while cur <= end:                       # one call per year keeps each response small
            nxt = min(date(cur.year, 12, 31), end)
            rows = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": cur.isoformat(), "to": nxt.isoformat(),
                                           "pointDirection": f"{opk}{pk}entry", "limit": -1}).get("operationalData", [])
            for r in rows:
                try:
                    rec[pd.Timestamp(r["periodFrom"][:10])] = float(r["value"]) / 1e6      # kWh/d -> GWh/d
                except (TypeError, ValueError, KeyError):
                    pass
            cur = nxt + timedelta(days=1)
        if rec:
            s = pd.Series(rec, dtype=float)
            series[name] = series[name].add(s, fill_value=0) if name in series else s
            print(f"  {name}: {len(rec)} days, {s.sum() / 1000:.1f} TWh", flush=True)
    new = pd.DataFrame(series).sort_index()
    if new.empty:
        print("no data")
        return
    comb = new.combine_first(old) if len(old) else new
    comb.loc[new.index, new.columns] = new
    comb = comb[sorted(comb.columns)].sort_index().round(3)
    comb.index.name = "date"
    print("TWh per year:\n" + (comb.groupby(comb.index.year).sum() / 1000).round(1).T.to_string())
    lines = ["Europe - ENTSOG entry points on the Norwegian border (physical flow)", "",
             "Source", "ENTSOG Transparency Platform, operational data: Physical Flow, daily (https://transparency.entsog.eu/).",
             "", "Units and definitions",
             "GWh per gas day. One column per entry point whose far side is Norway, named <country>_<point>. Emden (EPT1, EPT2, NPT) publishes "
             "no data under any indicator, so Germany's Norwegian volume here is Dornum only; the Gassco flow to Germany is the full figure. "
             "UK Easington and St Fergus mix UK North Sea gas with Norwegian gas.",
             f"Re-fetches the last {RELOAD_DAYS} days each run; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} days, {comb.index.min():%Y-%m-%d} to {comb.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
