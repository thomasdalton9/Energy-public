"""
Monthly natural gas supply and consumption by country from Eurostat (nrg_cb_gasm), the official statistics benchmark for
the ENTSOG-based gas balances:

  output/Data and Chart Outputs/eurostat_gas_monthly.xlsx
    sheet "Monthly": month, then <CC>_<item> in GWh (gross calorific value) for every EU27 country (+ NO, UK where reported):
        IC   inland consumption (observed)            FC_IND  final consumption, industry
        FC_OTH  final consumption, other sectors      PWR     input to electricity and heat generation (main producers)
        IMP  imports   EXP  exports   PROD  indigenous production   STK  stock change (negative = withdrawal)
    sheet "Units": source and definitions

Source: Eurostat dissemination API, dataset nrg_cb_gasm (https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_gasm).
Free, no key. TJ (GCV) is converted to GWh (/3.6). Published about a month after the reference month and revised.

Incremental: the committed workbook is the history store; each run re-fetches the last 12 months (Eurostat revises) plus any
missing months; history from 2021-01.

Usage: python3 EUROSTAT_GAS_MONTHLY.py [--out-dir DIR] [--start 2021-01]
"""
import argparse
import math
import os
import sys
import time
from datetime import datetime, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "eurostat_gas_monthly.xlsx"
URL = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm"
ITEMS = {"IC_OBS": "IC", "FC_IND": "FC_IND", "FC_OTH": "FC_OTH", "TI_EHG_MAP": "PWR", "IMP": "IMP", "EXP": "EXP", "IPRD": "PROD",
         "STK_CHG_MG": "STK"}
REVISION_MONTHS = 12


def fetch(since):
    params = [("format", "JSON"), ("lang", "EN"), ("unit", "TJ_GCV"), ("siec", "G3000"), ("sinceTimePeriod", since)] + \
             [("nrg_bal", k) for k in ITEMS]
    for i in range(4):
        r = requests.get(URL, params=params, timeout=240)
        if r.ok:
            return r.json()
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(5 * (i + 1))
            continue
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
    raise RuntimeError("gave up")


def to_frame(j):
    dims, size = j["id"], j["size"]
    idx = {d: list(j["dimension"][d]["category"]["index"]) for d in dims}
    strides = [math.prod(size[i + 1:]) for i in range(len(size))]
    rec = {}
    for k, v in j["value"].items():
        k = int(k)
        c = {d: idx[d][(k // strides[i]) % size[i]] for i, d in enumerate(dims)}
        item = ITEMS.get(c["nrg_bal"])
        if item is None or c["geo"] in ("EU27_2020", "EA20", "EU28", "EA19"):
            continue
        rec.setdefault(pd.Timestamp(c["time"] + "-01"), {})[f"{c['geo']}_{item}"] = float(v) / 3.6    # TJ -> GWh
    return pd.DataFrame.from_dict(rec, orient="index").sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    old = pd.DataFrame()
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Monthly", index_col=0)
            old.index = pd.to_datetime(old.index)
        except Exception as e:  # noqa: BLE001
            print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
    since = args.start
    if len(old):
        since = max(args.start, (old.index.max() - pd.DateOffset(months=REVISION_MONTHS)).strftime("%Y-%m"))
    new = to_frame(fetch(since))
    print(f"fetched {len(new)} months {new.index.min():%Y-%m}..{new.index.max():%Y-%m}, {new.shape[1]} series (from {since})")
    comb = new.combine_first(old) if len(old) else new
    comb.loc[new.index, new.columns] = new          # revisions overwrite
    comb = comb.sort_index().round(1)
    comb.index.name = "month"
    ic = comb[[c for c in comb if c.endswith("_IC")]]
    ann = (ic.groupby(ic.index.year).sum(min_count=12) / 1000).round(0)
    ann.columns = [c[:-3] for c in ann.columns]
    print("inland consumption, TWh per full year:\n" + ann.T.to_string())
    lines = ["Europe - monthly natural gas balance by country (Eurostat nrg_cb_gasm)", "",
             "Source", "Eurostat, dataset nrg_cb_gasm (https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_gasm). Free, no key.",
             "", "Units and definitions",
             "GWh per month, gross calorific value (Eurostat TJ / 3.6). <CC>_IC inland consumption (observed); _FC_IND final consumption "
             "industry; _FC_OTH final consumption other sectors (households, services); _PWR input to electricity and heat generation "
             "(main-activity producers); _IMP imports; _EXP exports; _PROD indigenous production; _STK stock change (negative = withdrawal). "
             "Published about a month after the reference month; recent months are revised.",
             f"Re-fetches the last {REVISION_MONTHS} months each run plus gaps; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} months, "
             f"{comb.index.min():%Y-%m} to {comb.index.max():%Y-%m}"]
    xlsx_notes.write_workbook(path, {"Monthly": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
