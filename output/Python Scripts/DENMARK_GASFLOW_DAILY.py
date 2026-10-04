"""
Denmark's daily commercial gas balance from Energinet's Gasflow dataset (Energi Data Service, free, no key):

  output/Data and Chart Outputs/denmark_energinet_gasflow_daily.xlsx
    sheet "Daily": date (gas day), GWh per day. Sign as published: positive = gas entering the Danish system,
    negative = gas leaving it.
        DK_from_north_sea     KWhFromNorthSea  gas entering from the North Sea (Danish fields, and Norwegian gas arriving via the
                                               Danish offshore pipelines that feeds Baltic Pipe)
        DK_from_tyra          kWhFromTyra      Tyra hub entry
        DK_biogas             KWhFromBiogas    biomethane injected into the Danish grids
        DK_to_denmark         KWhToDenmark     gas delivered to Danish consumers from the transmission system (negative)
        DK_storage            KWhToOrFromStorage  + withdrawal from / - injection into Gas Storage Denmark
        DK_germany            KWhToOrFromGermany  + imports from / - exports to Germany
        DK_to_sweden          KWhToSweden      exports to Sweden (negative)
        DK_to_poland          KWhToPoland      exports to Poland through Baltic Pipe (negative)
    sheet "Units": source and definitions

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days (the operator restates recent days)
plus any gap; history from 2021-01-01.

Usage: python3 DENMARK_GASFLOW_DAILY.py [--out-dir DIR] [--start 2021-01-01]
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
FILE = "denmark_energinet_gasflow_daily.xlsx"
URL = "https://api.energidataservice.dk/dataset/Gasflow"
REVISION_DAYS = 14
FIELDS = {"KWhFromNorthSea": "DK_from_north_sea", "kWhFromTyra": "DK_from_tyra", "KWhFromBiogas": "DK_biogas",
          "KWhToDenmark": "DK_to_denmark", "KWhToOrFromStorage": "DK_storage", "KWhToOrFromGermany": "DK_germany",
          "KWhToSweden": "DK_to_sweden", "KWhToPoland": "DK_to_poland"}
COLUMNS = list(FIELDS.values())


def fetch(d0, d1):
    for i in range(4):
        try:
            r = requests.get(URL, params={"start": d0.isoformat(), "end": (d1 + timedelta(days=1)).isoformat(), "limit": 100000,
                                          "sort": "GasDay ASC"}, timeout=90)
            if r.ok:
                return r.json().get("records", [])
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError("gave up on Energinet")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
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
    if fs > start and not set(COLUMNS) <= set(old.columns):
        fs = start
    rows = fetch(fs, today)
    d = pd.DataFrame(rows)
    if d.empty:
        print("no rows")
        return
    d["date"] = pd.to_datetime(d["GasDay"]).dt.normalize()
    d = d.drop_duplicates("date", keep="last").set_index("date").sort_index()
    new = pd.DataFrame({new_c: pd.to_numeric(d[src], errors="coerce") / 1e6 for src, new_c in FIELDS.items() if src in d})
    for c in COLUMNS:
        if c not in new:
            new[c] = float("nan")
    comb = new[COLUMNS].combine_first(old[[c for c in COLUMNS if c in old]]) if len(old) else new[COLUMNS]
    comb.loc[new.index, COLUMNS] = new[COLUMNS]
    comb = comb.sort_index().round(3)
    comb.index.name = "date"
    print(f"fetched {len(new)} days from {fs}; workbook {comb.index.min():%Y-%m-%d} .. {comb.index.max():%Y-%m-%d}, {len(comb)} days")
    ann = (comb.groupby(comb.index.year).sum() / 1000).round(1)
    print("TWh per year:\n" + ann.T.to_string())
    lines = ["Denmark - daily commercial gas balance (Energinet Gasflow)", "",
             "Source", "Energinet, Energi Data Service dataset Gasflow (https://www.energidataservice.dk/tso-gas/Gasflow). Free, no key.",
             "", "Units and definitions",
             "GWh per gas day (the source is kWh). Positive = gas entering the Danish system, negative = leaving. DK_from_north_sea = North Sea "
             "entry (Danish fields plus Norwegian gas arriving through the offshore pipelines that supplies Baltic Pipe); DK_from_tyra = Tyra "
             "hub entry; DK_biogas = biomethane injected; DK_to_denmark = delivered to Danish consumers (negative); DK_storage = withdrawal (+) "
             "or injection (-) at Gas Storage Denmark; DK_germany = imports (+) / exports (-) at the German border; DK_to_sweden and "
             "DK_to_poland = exports. The columns net to about zero each day (linepack and allocation differences remain).",
             f"Re-fetches the last {REVISION_DAYS} days each run; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} days, "
             f"{comb.index.min():%Y-%m-%d} to {comb.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
