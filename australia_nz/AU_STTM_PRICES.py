"""
Australia east coast gas hub prices: AEMO Short Term Trading Market (STTM)
ex-ante market price, A$/GJ, daily, for the Sydney, Adelaide and Brisbane
hubs (public NEMWEB report INT651, no key):

  https://nemweb.com.au/Reports/CURRENT/STTM/int651_v1_ex_ante_market_price_rpt_1.csv

The report holds a rolling window of recent gas days; each run merges it
into the saved sheet, so the history builds up from the first run (and any
older days the report still carries).

Writes au_gas_prices.xlsx, sheet "Daily": date, Sydney, Adelaide, Brisbane (A$/GJ).

Usage: python3 AU_STTM_PRICES.py [--out "output/Data and Chart Outputs/au_gas_prices.xlsx"]
"""
import argparse
import io
import os
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://nemweb.com.au/Reports/CURRENT/STTM/int651_v1_ex_ante_market_price_rpt_1.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_gas_prices.xlsx")
HUBS = ["Sydney", "Adelaide", "Brisbane"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    r = requests.get(URL, headers=HEADERS, timeout=(10, 120))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    d["date"] = pd.to_datetime(d["gas_date"], format="%d %b %Y", errors="coerce")
    d["price"] = pd.to_numeric(d["ex_ante_market_price"], errors="coerce")
    d = d.dropna(subset=["date"]).sort_values("approval_datetime")
    new = d.pivot_table(index="date", columns="hub_name", values="price", aggfunc="last")
    new = new.reindex(columns=[h for h in HUBS if h in new.columns] + [h for h in new.columns if h not in HUBS])
    print(f"INT651: {len(d)} rows, {len(new)} gas days {new.index.min():%Y-%m-%d}..{new.index.max():%Y-%m-%d}",
          flush=True)
    try:
        old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
        old.index = pd.to_datetime(old.index, errors="coerce")
        new = new.combine_first(old[old.index.notna()])[new.columns]
    except (FileNotFoundError, ValueError, KeyError, OSError):
        pass
    new.index.name = "date"
    print(new.tail(3).to_string(), flush=True)
    notes = [
        "UNITS",
        "A$ per GJ, STTM ex-ante market price for each gas day (the day-ahead clearing price at each hub; "
        "gas day 06:30-06:30 AEST).",
        "",
        "COVERAGE",
        f"Sydney, Adelaide and Brisbane STTM hubs, daily, {new.index.min():%d %b %Y} to {new.index.max():%d %b %Y}. "
        "History builds from the first run (the report is a rolling window). The Victorian DWGM price is not "
        "included yet.",
        "",
        "SOURCE",
        f"AEMO Short Term Trading Market, report INT651 (ex-ante market price): {URL}",
        "https://aemo.com.au/energy-systems/gas/short-term-trading-market-sttm",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": new.round(4)}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
