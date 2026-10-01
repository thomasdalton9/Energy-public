"""
Pull the daily Henry Hub Natural Gas Spot Price and maintain a growing
archive.

Data source: FRED (Federal Reserve Economic Data), series DHHNGSP -
https://fred.stlouisfed.org/graph/fredgraph.csv?id=DHHNGSP - a plain CSV
download, genuinely open, no API key or account needed (unlike getting
the same underlying EIA data via api.eia.gov, which needs a free EIA
key). Found via HENRY_HUB_FRED_INSPECT.py / HENRY_HUB_QUICK_PROBE.py.

Unlike MISO's real-time API (only "today"/"yesterday", no historical
range), this single request returns FRED's entire history for the
series in one response - back to 1997-01-07 as of the discovery pull.
So this one script is both the daily updater and the one-time
historical backfill: the very first run seeds the whole archive, and
every run after just upserts whatever's new (usually one row, since
Henry Hub spot prices publish once per business day with roughly a
one-day lag - no weekend/holiday rows, so gaps in the date index are
expected and not a sign of missing data).

Price is USD per MMBtu. FRED marks a missing/not-yet-published
observation as "." in the CSV rather than omitting the row - those are
dropped rather than kept as a fabricated value.
"""

import argparse
import os
import sys
from datetime import date

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DHHNGSP"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)

# FRED publishes Henry Hub spot price with roughly a one-day lag and
# only on business days; a few extra days of slack absorb a holiday
# weekend before this is treated as a genuinely stale source.
STALE_AFTER_DAYS = 7

DEFAULT_OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "henry_hub_daily.xlsx"
)


FETCH_ATTEMPTS = 5
RETRY_BACKOFF_SECONDS = [5, 15, 30, 60]  # between attempts 1-2, 2-3, 3-4, 4-5


def fetch_series():
    """FRED's CSV endpoint is occasionally slow enough to exceed a
    30s read timeout (observed as fast as 0.24s and, minutes later,
    timing out three times in a row) - looks like transient backend
    load rather than a real block, so retry with backoff before giving
    up."""
    import time

    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            last_error = e
            print(f"  attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                wait = RETRY_BACKOFF_SECONDS[attempt - 1]
                print(f"  waiting {wait}s before retrying...", file=sys.stderr)
                time.sleep(wait)
    raise last_error


def parse_series(csv_text):
    from io import StringIO

    df = pd.read_csv(StringIO(csv_text))
    date_col, value_col = df.columns[0], df.columns[1]
    df[date_col] = pd.to_datetime(df[date_col]).dt.date
    df["Henry_Hub_USD_per_MMBtu"] = pd.to_numeric(df[value_col], errors="coerce")
    df = df.dropna(subset=["Henry_Hub_USD_per_MMBtu"])
    df = df.set_index(date_col)[["Henry_Hub_USD_per_MMBtu"]]
    df.index.name = "date"
    return df.sort_index()


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new_df):
    if existing.empty:
        combined = new_df
    else:
        combined = pd.concat([existing, new_df])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


def format_date_column(path, sheet="Data"):
    import openpyxl

    wb = openpyxl.load_workbook(path)
    if sheet not in wb.sheetnames:
        return
    ws = wb[sheet]
    for (cell,) in ws.iter_rows(min_row=2, max_col=1):
        cell.number_format = "dd-mmm-yyyy"
    ws.column_dimensions["A"].width = 14
    wb.save(path)


NOTES_LINES = [
    "UNITS",
    "Henry_Hub_USD_per_MMBtu: US dollars per million British thermal units (MMBtu) - the standard "
    "Henry Hub natural gas spot price quote.",
    "",
    "COVERAGE",
    "Published on business days only (no weekend/holiday rows) - gaps in the date index are expected, "
    "not missing data. FRED publishes with roughly a one-day lag from the trade date.",
    "",
    "TIMESTAMPS",
    "The 'date' index is the trade date FRED assigns the observation to, as published - not converted "
    "or shifted.",
    "",
    "SOURCE",
    f"FRED (Federal Reserve Economic Data), series DHHNGSP: {URL} - sourced from EIA, but this CSV "
    "endpoint is public with no API key needed (unlike EIA's own api.eia.gov for the same data).",
]
NOTES_SECTION_TITLES = {"UNITS", "COVERAGE", "TIMESTAMPS", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    print(f"Fetching {URL} ...", file=sys.stderr)
    csv_text = fetch_series()
    new_df = parse_series(csv_text)
    print(f"  {len(new_df)} published observations, {new_df.index.min()} to {new_df.index.max()}",
          file=sys.stderr)

    existing = load_archive(args.out)
    before_days = set(existing.index) if not existing.empty else set()
    combined = upsert(existing, new_df)
    new_days = sorted(set(combined.index) - before_days)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_column(args.out)

    print(f"Added {len(new_days)} new day(s){' - ' + str(new_days[-5:]) if new_days else ''}")
    print(f"Archive now has {len(combined)} days ({combined.index.min()} to {combined.index.max()}). "
          f"Saved to {args.out}")
    print(combined.tail())

    latest = max(combined.index) if not combined.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest saved day is {latest} ({age} days old) - "
              f"{URL} may have changed or stopped updating.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
