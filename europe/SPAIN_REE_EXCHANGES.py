"""
Spain's physical electricity exchanges with Morocco and Andorra (Red Electrica REData, apidatos.ree.es, daily), one workbook:

  output/Data and Chart Outputs/spain_ree_exchanges_daily.xlsx
    sheet "Daily": date, Morocco_MWh, Andorra_MWh (net import, positive = into Spain, negative = export), MA_AD_NetImports_MWh (sum), plus
                   France_MWh and Portugal_MWh (REE's own figures, kept only to check ENTSO-E's)
    sheet "Units": source and definitions

Why: ENTSO-E has no Morocco or Andorra bidding zones, so Spain's ENTSO-E net imports (France + Portugal only) miss the exports to Morocco (1.9 TWh in
2023, 2.5 in 2024, 3.8 in 2025) and Andorra (0.2-0.24 TWh a year). That was Spain's whole 'supply above load' (1.0-1.1%): REE's four border balances add up
exactly to its published cross-border balance (13.96 / 10.23 / 12.80 TWh net export 2023-25), and its France and Portugal figures match ENTSO-E's within
0.1 TWh. The Europe master adds Morocco + Andorra to Spain's net imports.

Incremental: the committed workbook is the history store; each run re-reads the last 30 days and fetches any later days.
Usage: python3 SPAIN_REE_EXCHANGES.py [--out-dir DIR] [--start 2021-01-01]
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

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "spain_ree_exchanges_daily.xlsx"
URL = "https://apidatos.ree.es/en/datos/intercambios/{}-frontera"
BORDERS = {"marruecos": "Morocco_MWh", "andorra": "Andorra_MWh", "francia": "France_MWh", "portugal": "Portugal_MWh"}
REVISION_DAYS = 30
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def fetch(border, d0, d1):
    """{date: net import MWh} (REE 'saldo': negative = export), one request per half-year."""
    out = {}
    t = d0
    while t < d1:
        t2 = min(d1, t + timedelta(days=180))
        params = {"start_date": f"{t:%Y-%m-%d}T00:00", "end_date": f"{t2 - timedelta(days=1):%Y-%m-%d}T23:59", "time_trunc": "day"}
        for attempt in range(5):
            try:
                r = requests.get(URL.format(border), params=params, headers=UA, timeout=(10, 60))
                if r.status_code == 200:
                    break
            except requests.RequestException:
                pass
            time.sleep(5 * (attempt + 1))
        else:
            raise RuntimeError(f"REData {border} {t:%Y-%m-%d} failed")
        for grp in r.json().get("included", []):
            if grp.get("type") == "saldo":
                for v in grp["attributes"].get("values", []):
                    out[date.fromisoformat(v["datetime"][:10])] = float(v["value"])
        time.sleep(1)
        t = t2
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    old = pd.DataFrame()
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.set_index("date").sort_index()
        except Exception:  # noqa: BLE001
            old = pd.DataFrame()
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    if len(old):
        start = max(start, old.index.max().date() - timedelta(days=REVISION_DAYS))
    end = datetime.now(timezone.utc).date()   # up to yesterday: today is incomplete
    new = pd.DataFrame({col: pd.Series(fetch(b, start, end)) for b, col in BORDERS.items()})
    new.index = pd.to_datetime(new.index)
    print(f"fetched {len(new)} days from {start}", flush=True)
    daily = pd.concat([old[~old.index.isin(new.index)], new]).sort_index() if len(old) else new.sort_index()
    daily = daily.dropna(subset=["Morocco_MWh", "Andorra_MWh"])
    daily["MA_AD_NetImports_MWh"] = daily["Morocco_MWh"] + daily["Andorra_MWh"]
    daily = daily[["MA_AD_NetImports_MWh", "Morocco_MWh", "Andorra_MWh", "France_MWh", "Portugal_MWh"]].round(1)
    daily.index.name = "date"
    print("TWh per year:\n" + (daily.groupby(daily.index.year).sum() / 1e6).round(3).T.to_string(), flush=True)
    lines = ["Spain - physical electricity exchanges with Morocco and Andorra (Red Electrica de Espana, REData)", "",
             "Source", "Red Electrica de Espana REData API, datos/intercambios/<border>-frontera (physical exchanges, daily), https://www.ree.es/en/datos/intercambios . Free, no key.",
             "", "Units and definitions",
             "MWh per day (REE local day), net import into Spain: positive = import, negative = export (REE 'saldo' = imports + exports). MA_AD_NetImports_MWh = "
             "Morocco + Andorra, the two Spanish borders that ENTSO-E does not publish (no bidding zone); the Europe master adds it to Spain's ENTSO-E net imports "
             "(France + Portugal). France_MWh and Portugal_MWh are REE's own figures for the other two borders and are kept only as a check on ENTSO-E's. "
             f"Re-reads the last {REVISION_DAYS} days each run.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": daily}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
