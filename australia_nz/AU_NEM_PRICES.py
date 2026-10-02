"""
Australia NEM wholesale electricity prices by region (state), from AEMO's
5-minute dispatch prices (DISPATCHPRICE, RRP in A$/MWh - public NEMWEB /
MMSDM data via nemosis, no key).

Writes au_nem_prices.xlsx:
  Daily average   A$/MWh, simple average of the day's 5-minute prices, per state
  Negative hours  hours per day with a negative price, per state
  Daily max       highest 5-minute price of the day, per state
  Daily min       lowest 5-minute price of the day, per state
Market day = intervals ending 00:05 to 24:00 (AEST). Intervention pricing runs are excluded.

Incremental: fetches months not saved yet plus the latest saved month (the month in progress); a new file
backfills from 2021, --max-months per run. Saves after every month.

Usage: python3 AU_NEM_PRICES.py [--out "output/Data and Chart Outputs/au_nem_prices.xlsx"]
"""
import argparse
import os
import shutil
import sys
from datetime import date

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

HISTORY_START = "2021-01-01"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_nem_prices.xlsx")
CACHE = os.path.join(ROOT, "nemosis_cache_prices")
STATES = ["NSW", "QLD", "VIC", "SA", "TAS"]
SHEETS = {"Daily average": "mean", "Negative hours": "neg", "Daily max": "max", "Daily min": "min"}


def fetch_month(start, end):
    from nemosis import dynamic_data_compiler

    os.makedirs(CACHE, exist_ok=True)
    try:
        d = dynamic_data_compiler(start.strftime("%Y/%m/%d %H:%M:%S"), end.strftime("%Y/%m/%d %H:%M:%S"),
                                  "DISPATCHPRICE", CACHE)
    finally:
        shutil.rmtree(CACHE, ignore_errors=True)
    if d is None or d.empty:
        return None
    if "INTERVENTION" in d:
        d = d[pd.to_numeric(d["INTERVENTION"], errors="coerce").fillna(0) == 0]
    d = d.assign(date=(pd.to_datetime(d["SETTLEMENTDATE"]) - pd.Timedelta(minutes=5)).dt.normalize(),
                 state=d["REGIONID"].str[:-1], rrp=pd.to_numeric(d["RRP"], errors="coerce"))
    d = d[(d["date"] >= start) & (d["date"] < end)]
    g = d.groupby(["date", "state"])["rrp"]
    out = {"mean": g.mean(), "max": g.max(), "min": g.min(),
           "neg": d.assign(neg=d["rrp"] < 0).groupby(["date", "state"])["neg"].sum() / 12.0}   # 5-min intervals -> hours
    return {k: v.unstack("state").reindex(columns=STATES) for k, v in out.items()}


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def save(path, frames):
    first = frames["mean"].dropna(how="all")
    notes = [
        "UNITS",
        "Daily average, Daily max, Daily min: A$/MWh, from AEMO's 5-minute regional reference prices (RRP). Daily "
        "average is the simple (time-weighted) average of the day's 288 prices, not demand-weighted.",
        "Negative hours: hours in the day with a negative 5-minute price (count of intervals / 12).",
        "Market day: intervals ending 00:05 to 24:00, AEST. Intervention pricing runs excluded.",
        "",
        "COVERAGE",
        f"NEM regions NSW (incl. ACT), QLD, VIC, SA, TAS, daily, {first.index.min():%d %b %Y} to "
        f"{first.index.max():%d %b %Y}. WA (WEM) is a separate market and not included.",
        "",
        "SOURCE",
        "AEMO NEMWEB / MMSDM DISPATCHPRICE (via nemosis)",
        "https://nemweb.com.au/",
    ]
    sheets = {}
    for name, key in SHEETS.items():
        f = frames[key].copy()
        f.index.name = "date"
        sheets[name] = f.round(2)
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-months", type=int, default=72)
    args = ap.parse_args()

    frames = {key: load(args.out, name) for name, key in SHEETS.items()}
    today = pd.Timestamp(date.today())
    months = pd.date_range(HISTORY_START, today, freq="MS")
    have = sorted(set(frames["mean"].index.to_period("M").to_timestamp())) if not frames["mean"].empty else []
    todo = sorted(set([m for m in months if m not in have] + have[-1:]))[:args.max_months]
    print(f"{len(have)} months saved; fetching {len(todo)}", flush=True)
    for m in todo:
        end = min(m + pd.offsets.MonthBegin(1), today)
        try:
            new = fetch_month(m, end)
        except Exception as e:  # noqa: BLE001 - keep the months already done
            print(f"  {m:%Y-%m}: failed {type(e).__name__}: {str(e)[:200]}", flush=True)
            continue
        if new is None:
            print(f"  {m:%Y-%m}: no rows", flush=True)
            continue
        for k in frames:
            old = frames[k]
            old = old[old.index.to_period("M") != m.to_period("M")] if not old.empty else old
            frames[k] = pd.concat([old, new[k]]).sort_index()
        print(f"  {m:%Y-%m}: avg " + ", ".join(f"{s} {new['mean'][s].mean():.0f}" for s in STATES if s in new["mean"]),
              flush=True)
        save(args.out, frames)
    if frames["mean"].empty:
        raise SystemExit("No NEM price data")
    print(f"Saved {args.out}: {len(frames['mean'])} days", flush=True)


if __name__ == "__main__":
    main()
