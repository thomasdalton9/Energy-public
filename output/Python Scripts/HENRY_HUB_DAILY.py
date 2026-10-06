"""
Pull the daily Henry Hub Natural Gas Spot Price and maintain a growing
archive (output/Data and Chart Outputs/henry_hub_daily.xlsx, sheet
'Data': date index + Henry_Hub_USD_per_MMBtu).

Source: EIA only, tried in order (first that answers wins):
  1. EIA API v2, series RNGWHHD (route natural-gas/pri/fut, daily) - needs
     the free EIA key, read from the EIA_API_KEY environment variable
     (repo secret EIA_API_KEY, as used by the EIA-930 workflow).
     Incremental: once the archive exists only the last two weeks are
     re-read.
  2. EIA's keyless history workbook
     https://www.eia.gov/dnav/ng/hist_xls/RNGWHHDd.xls (full history).
The previous FRED (DHHNGSP) source kept failing with read timeouts from
GitHub Actions (Sep 2026) and was removed on 2026-10-01.

History runs back to 1997-01-07. Henry Hub spot prices publish once per
business day with a lag of a few days - no weekend/holiday rows, so gaps
in the date index are expected and not a sign of missing data. Price is
USD per MMBtu; missing observations are dropped, never filled.
Charts: add_charts.py (registry entry henry_hub_daily.xlsx) after the pull.
"""

import argparse
import io
import os
import sys
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

EIA_API = "https://api.eia.gov/v2/natural-gas/pri/fut/data/"
EIA_SERIES = "RNGWHHD"
EIA_XLS = "https://www.eia.gov/dnav/ng/hist_xls/RNGWHHDd.xls"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
COL = "Henry_Hub_USD_per_MMBtu"
COL_FM = "Henry_Hub_front_month_USD_per_MMBtu"
YAHOO_FM = "https://{host}.finance.yahoo.com/v8/finance/chart/NG=F?period1={p1}&period2={p2}&interval=1d"
FM_SEED_START = date(2000, 9, 1)  # first full month of Yahoo's NG=F history

# EIA publishes Henry Hub spot price with a lag of a few business days;
# a few extra days of slack absorb a holiday weekend before this is
# treated as a genuinely stale source.
STALE_AFTER_DAYS = 10
OVERLAP_DAYS = 14  # incremental runs re-read the last two weeks (EIA occasionally revises recent days)

DEFAULT_OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "Data and Chart Outputs",
    "henry_hub_daily.xlsx"
)

FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [5, 20]


def get(url, **kw):
    """GET with a short retry loop (sources are tried in turn, so each one gives up fairly quickly)."""
    import time

    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last_error = e
            print(f"  attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS[attempt - 1])
    raise last_error


def tidy(dates, values):
    df = pd.DataFrame({"date": pd.DatetimeIndex(pd.to_datetime(list(dates), errors="coerce")).date,
                       COL: pd.to_numeric(list(values), errors="coerce")})
    df = df.dropna().set_index("date").sort_index()
    return df[~df.index.duplicated(keep="last")]


def fetch_eia_api(start=None):
    """EIA API v2, series RNGWHHD (needs EIA_API_KEY). Incremental from `start`; pages of 5000 rows."""
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise RuntimeError("EIA_API_KEY not set")
    rows, offset = [], 0
    while True:
        params = {"api_key": key, "frequency": "daily", "data[0]": "value", "facets[series][]": EIA_SERIES,
                  "sort[0][column]": "period", "sort[0][direction]": "asc", "offset": offset, "length": 5000}
        if start:
            params["start"] = start.isoformat()
        resp = get(EIA_API, params=params).json()["response"]
        data = resp.get("data", [])
        rows += data
        offset += len(data)
        if not data or offset >= int(resp.get("total", 0)):
            break
    units = {r.get("units") for r in rows}
    print(f"  EIA API: {len(rows)} rows, units {units}", file=sys.stderr)
    return tidy([r["period"] for r in rows], [r["value"] for r in rows])


def fetch_eia_xls(start=None):
    """EIA's keyless history workbook (sheet 'Data 1': Date, Henry Hub spot $/MMBtu, header on row 3)."""
    raw = pd.read_excel(io.BytesIO(get(EIA_XLS).content), sheet_name="Data 1", header=2)
    print(f"  EIA XLS columns: {list(raw.columns)}", file=sys.stderr)
    return tidy(raw.iloc[:, 0], raw.iloc[:, 1])


SOURCES = [("EIA API v2 (RNGWHHD)", fetch_eia_api), ("EIA history XLS (RNGWHHDd.xls)", fetch_eia_xls)]


def fetch_series(start=None):
    """First source that answers wins: EIA API, then EIA's keyless XLS of the same series."""
    for name, fn in SOURCES:
        print(f"Fetching from {name} ...", file=sys.stderr)
        try:
            df = fn(start)
        except Exception as e:  # noqa: BLE001 - fall through to the next source
            print(f"  {name} failed: {type(e).__name__}: {e}", file=sys.stderr)
            continue
        if not df.empty:
            return df, name
        print(f"  {name} returned no rows", file=sys.stderr)
    raise SystemExit("all Henry Hub sources failed")


def fetch_front_month(start=None):
    """NYMEX Henry Hub front-month futures daily close (USD/MMBtu) from Yahoo Finance's unofficial chart API, ticker NG=F
    (continuous front month: it jumps at each contract roll, unadjusted). EIA's own futures series stops on 2024-04-05."""
    import time
    from curl_cffi import requests as cr

    p1 = int(pd.Timestamp(start or FM_SEED_START).timestamp())
    p2 = int(time.time()) + 86400
    last_error = None
    for host in ("query1", "query2"):
        for attempt in range(1, FETCH_ATTEMPTS + 1):
            try:
                r = cr.get(YAHOO_FM.format(host=host, p1=p1, p2=p2), impersonate="chrome", timeout=40,
                           headers={"Accept": "application/json"})
                r.raise_for_status()
                res = r.json()["chart"]["result"][0]
                df = pd.DataFrame({"date": pd.to_datetime(res["timestamp"], unit="s").normalize().date,
                                   COL_FM: res["indicators"]["quote"][0]["close"]})
                df = df.dropna().set_index("date").sort_index()
                return df[~df.index.duplicated(keep="last")]
            except Exception as e:  # noqa: BLE001
                last_error = e
                print(f"  Yahoo NG=F {host} attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}",
                      file=sys.stderr)
                time.sleep(RETRY_BACKOFF_SECONDS[min(attempt - 1, 1)])
    raise last_error


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new_df):
    """Per-column upsert: days in new_df replace those days in existing for new_df's columns only."""
    if existing.empty:
        combined = new_df
    else:
        combined = existing.copy()
        for c in new_df.columns:
            if c not in combined.columns:
                combined[c] = float("nan")
        combined = combined.reindex(combined.index.union(new_df.index))
        for c in new_df.columns:
            combined.loc[new_df.index, c] = new_df[c]
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
    "Henry_Hub_front_month_USD_per_MMBtu: NYMEX Henry Hub natural gas futures, front-month contract, daily close, USD/MMBtu "
    "(Yahoo Finance NG=F continuous series, from 2000-09; trading days only, so it also has days with no spot row). It is "
    "unadjusted: the series steps when the front contract rolls to the next month.",
    "",
    "COVERAGE",
    "Daily from 1997-01-07. Published on business days only (no weekend/holiday rows) - gaps in the date "
    "index are expected, not missing data. EIA publishes with a lag of a few business days.",
    "",
    "TIMESTAMPS",
    "The 'date' index is the trade date EIA assigns the observation to, as published - not converted "
    "or shifted.",
    "",
    "SOURCE",
    f"EIA (US Energy Information Administration), Henry Hub Natural Gas Spot Price, series {EIA_SERIES}: "
    f"EIA API v2 ({EIA_API}, needs an API key), falling back to EIA's keyless history workbook {EIA_XLS}. "
    "Only days a source returns are upserted; existing days are never deleted.",
    "Front month: Yahoo Finance chart API (query1.finance.yahoo.com/v8/finance/chart/NG=F), unofficial and undocumented, "
    "used because EIA's futures series (RNGC1) stops on 2024-04-05; if Yahoo fails the run keeps the saved front-month "
    "history. Re-read: last 14 days.",
    "Source changed 2026-10-01: EIA is now the only source (the earlier third-party re-publication of this "
    "series kept timing out from GitHub Actions and was dropped).",
]
NOTES_SECTION_TITLES = {"UNITS", "COVERAGE", "TIMESTAMPS", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    existing = load_archive(args.out)
    # Incremental: only re-read the last couple of weeks once the archive holds the history.
    start = max(existing.index) - timedelta(days=OVERLAP_DAYS) if not existing.empty else None
    new_df, source = fetch_series(start)
    print(f"  {source}: {len(new_df)} observations, {new_df.index.min()} to {new_df.index.max()}",
          file=sys.stderr)

    before_days = set(existing.index) if not existing.empty else set()
    combined = upsert(existing, new_df)
    # Front-month futures (non-fatal: a Yahoo failure keeps the saved history)
    try:
        fm_start = None
        if COL_FM in combined.columns and combined[COL_FM].notna().any():
            fm_start = combined[COL_FM].dropna().index.max() - timedelta(days=OVERLAP_DAYS)
        fm = fetch_front_month(fm_start)
        combined = upsert(combined, fm)
        print(f"  Yahoo NG=F front month: {len(fm)} rows, {fm.index.min()} to {fm.index.max()}", file=sys.stderr)
    except Exception as e:  # noqa: BLE001
        print(f"  front-month pull failed, keeping saved history: {type(e).__name__}: {e}", file=sys.stderr)
    new_days = sorted(set(combined.index) - before_days)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES + ["", f"Last run's source: {source}"],
                              NOTES_SECTION_TITLES)
    format_date_column(args.out)

    print(f"Source used: {source}")
    print(f"Added {len(new_days)} new day(s){' - ' + str(new_days[-5:]) if new_days else ''}")
    print(f"Archive now has {len(combined)} days ({combined.index.min()} to {combined.index.max()}). "
          f"Saved to {args.out}")
    print(combined.tail())

    latest = max(combined[COL].dropna().index) if not combined.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest saved day is {latest} ({age} days old) - "
              f"{source} may have changed or stopped updating.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
