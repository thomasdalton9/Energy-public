"""
Day-ahead electricity prices for Europe from the ENTSO-E Transparency Platform (Day-ahead Prices [12.1.D], API
document A44), one workbook:

  output/Data and Chart Outputs/europe_power_prices_daily.xlsx
    sheet "Daily": date, then one column per country (EUR/MWh, mean of the day's day-ahead prices)
    sheet "Units": source and definitions

One bidding zone per country: the country's only zone, or for multi-zone countries Denmark DK1, Sweden SE3,
Norway NO2 and Italy IT-North (Italy's single national price, PUN, is set by the GME and is not on ENTSO-E).
Columns are named "<country>" or "<country> (<zone>)".

A day is the UTC day; the mean is over the finest resolution present (hourly, or 15-minute since the
15-minute day-ahead market started) and a day with an incomplete price set is left out and fetched again later.

Incremental: reads the committed workbook and re-fetches the last 7 days plus any gap in the last 120 days;
backfill starts 2021-01-01 in 60-day windows.

Usage: python3 ENTSOE_PRICES_DAILY.py [--out-dir DIR] [--start 2021-01-01]
Requires ENTSOE_API_KEY (environment / GitHub secret, or api_keys.py).
"""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
import entsoe_common as C  # noqa: E402
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_power_prices_daily.xlsx"
REVISION_DAYS = 7
GAP_DAYS = 120
WINDOW_DAYS = 60
ZONE_PREF = {"DK": "DK1", "SE": "SE3", "NO": "NO2", "IT": "IT-North"}
OK_COUNTS = set(range(22, 27)) | set(range(88, 101))   # hourly (23-25) or 15-minute (92-100) day, with slack


def columns():
    """[(column name, country code, zone label, EIC)] one zone per country."""
    out = []
    for code, (name, slug, zones) in C.COUNTRIES.items():
        label, eic = next(((l, e) for l, e in zones if l == ZONE_PREF.get(code)), zones[0])
        out.append((f"{name} ({label})" if len(zones) > 1 else name, code, label, eic))
    return out


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    return d.dropna(subset=["date"]).set_index("date").sort_index()


def zone_prices(eic, d0, d1, deadline):
    """{date: mean EUR/MWh} for complete days of one zone in [d0, d1)."""
    out = {}
    for w0, w1 in C.windows(d0, d1, WINDOW_DAYS):
        if time.time() > deadline:
            raise TimeoutError
        res = C.fetch_split({"documentType": "A44", "in_Domain": eic, "out_Domain": eic}, w0, w1,
                            C.parse_prices, C.merge_prices)
        time.sleep(0.25)
        for d, vals in (res or {}).items():
            if len(vals) in OK_COUNTS and w0.date() <= d < w1.date():
                out[d] = sum(vals) / len(vals)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    ap.add_argument("--max-minutes", type=float, default=120.0)
    args = ap.parse_args()
    C.api_key()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    deadline = time.time() + args.max_minutes * 60
    start = datetime.strptime(args.start, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    end = C.today_utc()
    old = read_existing(path)
    cols = columns()
    for col, *_ in cols:
        if col not in old:
            old[col] = float("nan")
    old = old[[c for c, *_ in cols]] if len(old) else pd.DataFrame(columns=[c for c, *_ in cols])
    old.index.name = "date"
    results, summary = {}, []
    for col, code, label, eic in cols:
        s = old[col].dropna() if len(old) else pd.Series(dtype=float)
        fs = start
        if len(s):
            fs = max(start, s.index.max().to_pydatetime().replace(tzinfo=timezone.utc) - timedelta(days=REVISION_DAYS))
            cutoff = (end - timedelta(days=GAP_DAYS)).date()
            gaps = [start + timedelta(days=i) for i in range((end - start).days)]
            gaps = [g for g in gaps if g.date() >= cutoff and pd.Timestamp(g.date()) not in s.index and g < fs]
            if gaps:
                fs = min(fs, gaps[0])
        try:
            got = zone_prices(eic, fs, end, deadline)
        except TimeoutError:
            print("time budget reached; stopping (columns done so far are saved)")
            break
        except Exception as e:  # noqa: BLE001
            print(f"{col}: FAILED {type(e).__name__}: {e}")
            summary.append((col, "FAILED", len(s)))
            continue
        results[col] = pd.Series({pd.Timestamp(d): v for d, v in got.items()}, dtype=float)
        n_total = len(set(s.index) | set(results[col].index))
        print(f"{col}: {len(got)} days fetched from {fs:%Y-%m-%d}; {n_total} saved", flush=True)
        summary.append((col, f"{n_total} days", n_total))
    new = pd.DataFrame(results)
    combined = old.copy() if len(old) else pd.DataFrame(columns=[c for c, *_ in cols])
    if not new.empty:
        combined = combined.reindex(combined.index.union(new.index))
        for col in new:
            combined.loc[new[col].dropna().index, col] = new[col].dropna()
    combined = combined.dropna(how="all").sort_index().round(2)
    combined.index.name = "date"
    if combined.empty:
        print("no price data")
        return
    lines = ["Europe - day-ahead electricity prices (ENTSO-E Transparency Platform)", "",
             "Source", "ENTSO-E Transparency Platform, Day-ahead Prices [12.1.D] (API document A44). https://transparency.entsoe.eu/",
             "", "Zones", "; ".join(f"{c}: {l} ({e})" for c, _, l, e in cols),
             "", "Units and definitions",
             "EUR/MWh, mean of the day's day-ahead prices for the UTC day (finest resolution present: hourly, or 15-minute "
             "since the 15-minute day-ahead market started). Days with an incomplete price set are left out and fetched "
             "again on a later run.",
             "One zone per country; for multi-zone countries Denmark DK1, Sweden SE3, Norway NO2, Italy IT-North.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps within {GAP_DAYS} days; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Zones", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(combined)} days x {combined.shape[1]} columns")
    print("\nSUMMARY")
    for col, msg, n in summary:
        print(f"  {col:40s} {msg}")


if __name__ == "__main__":
    main()
