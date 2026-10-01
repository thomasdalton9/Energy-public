"""
Pull Ireland/Northern Ireland's (all-island) system demand, wind
generation, and CO2 intensity from EirGrid/SONI's Smart Grid Dashboard
and maintain a growing archive.

Data source: https://www.smartgriddashboard.com/api/chart/ - found via
Playwright network capture of the dashboard's own homepage
(IRELAND_TURKEY_POWER_PLAYWRIGHT_DISCOVERY.py), no login or key needed.
Confirmed live (IRELAND_SMARTGRID_API_INSPECT.py /
IRELAND_SMARTGRID_RANGE_INSPECT.py):
  - chartType=demand (SYSTEM_DEMAND + DEMAND_FORECAST_VALUE), wind
    (WIND_ACTUAL + WIND_FCAST), co2 (CO2_INTENSITY) all return real
    15-minute-interval data, confirmed back to at least 2015-01-15.
  - dateRange="day" (NOT "week"/"month"/"year", which give inconsistent
    or empty results) accepts an arbitrary dateFrom/dateTo SPAN, not
    just a single day - confirmed a full 31-day span returns all 31
    days (2,976 rows = 31 * 96) in ONE request. So, unlike Mexico's
    CENACE (which needs one request per day), this backfills in large
    date-range chunks.
  - chartType=generation/areas=fuelmix exists but only ever returns ~5
    rows (one snapshot per broad category - FUEL_COAL, FUEL_RENEW, etc,
    as of a single recent timestamp, not a real generation-by-fuel time
    series) - not used here; a genuine hourly fuel-mix series was not
    found via this endpoint.

All values are for "region=ALL" (all-island, ROI + NI combined) - the
dashboard also supports "ROI" and "NI" separately, not pulled here.
"""

import argparse
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://www.smartgriddashboard.com/api/chart/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://www.smartgriddashboard.com/",
}
TIMEOUT = (10, 30)
FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [5, 15]

CHART_TYPES = {
    "demand": "demandactual,demandforecast",
    "wind": "windactual,windforecast",
    "co2": "co2intensity",
}
FIELD_RENAME = {
    "SYSTEM_DEMAND": "Demand_Actual_MW",
    "DEMAND_FORECAST_VALUE": "Demand_Forecast_MW",
    "WIND_ACTUAL": "Wind_Actual_MW",
    "WIND_FCAST": "Wind_Forecast_MW",
    "CO2_INTENSITY": "CO2_Intensity_gCO2_per_kWh",
}

DATA_START = date(2015, 1, 15)  # confirmed live; may go back further, not yet tested
# The API silently returns 0 rows (HTTP 200, empty "Rows") past some
# width threshold - confirmed live (IRELAND_SMARTGRID_BUG_ISOLATE.py)
# that a 90-day span always fails (regardless of year or how many
# "areas" are requested) while a 31-day span always works. The real
# cutoff was never pinned down more precisely than "somewhere between
# 31 and 90" - 30 days is safely inside the confirmed-working side.
CHUNK_DAYS = 30

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "ireland_smartgrid_15min.xlsx")


def fetch_chunk(chart_type, areas, start, end):
    params = {
        "region": "ALL",
        "chartType": chart_type,
        "dateRange": "day",  # NOT week/month/year - see module docstring
        "dateFrom": start.strftime("%d-%b-%Y"),
        "dateTo": end.strftime("%d-%b-%Y"),
        "areas": areas,
    }
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(URL, headers=HEADERS, params=params, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json().get("Rows", [])
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}",
                  file=sys.stderr, flush=True)
            if attempt < FETCH_ATTEMPTS:
                wait = RETRY_BACKOFF_SECONDS[attempt - 1]
                print(f"    waiting {wait}s before retrying...", file=sys.stderr, flush=True)
                time.sleep(wait)
    raise last_error


CHART_TYPE_COLUMNS = {
    "demand": ["Demand_Actual_MW", "Demand_Forecast_MW"],
    "wind": ["Wind_Actual_MW", "Wind_Forecast_MW"],
    "co2": ["CO2_Intensity_gCO2_per_kWh"],
}
# Re-checkpoint to disk every this-many chunks within a single chart
# type's fetch, not just once per chart type - a full initial backfill
# is ~140 chunks per type (~4,270 days / 30), and without this a
# workflow timeout mid-fetch would discard everything fetched so far
# (confirmed live: the script previously only wrote xlsx_notes.write_
# workbook() once, at the very end of ALL three chart types combined -
# a 15-minute internal `timeout 900` wrapper killing that run would
# have lost the entire backfill with nothing committed).
CHECKPOINT_EVERY_N_CHUNKS = 10


def rows_to_wide(rows):
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["datetime"] = pd.to_datetime(df["EffectiveTime"], format="%d-%b-%Y %H:%M:%S")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    df["column"] = df["FieldName"].map(lambda f: FIELD_RENAME.get(f, f))
    wide = df.pivot_table(index="datetime", columns="column", values="Value", aggfunc="last")
    wide.index.name = "datetime"
    ordered = [c for c in FIELD_RENAME.values() if c in wide.columns]
    other = [c for c in wide.columns if c not in ordered]
    return wide[ordered + other].sort_index()


def merge_columns(existing, new_wide):
    """Upsert new_wide's columns/rows into existing. new_wide's values
    win wherever it has data (even re-fetched overlap); everything else
    in existing (other chart types' columns, other rows) is preserved."""
    if new_wide.empty:
        return existing
    if existing.empty:
        combined = new_wide
    else:
        combined = new_wide.combine_first(existing)
    combined.index.name = "datetime"
    return combined.sort_index()


def resume_date_for(existing, chart_type):
    cols = [c for c in CHART_TYPE_COLUMNS[chart_type] if c in existing.columns]
    if not cols:
        return None
    sub = existing[cols].dropna(how="all")
    if sub.empty:
        return None
    return sub.index.max().date()


def fetch_and_merge_chart_type(chart_type, areas, start, end, existing, save_callback):
    """Fetches one chart type's full [start, end] range in CHUNK_DAYS
    chunks, checkpointing (merging into existing + saving to disk) every
    CHECKPOINT_EVERY_N_CHUNKS chunks so a mid-fetch interruption keeps
    whatever was already fetched instead of losing it all."""
    rows = []
    chunk_start = start
    chunk_count = 0
    while chunk_start <= end:
        chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS - 1), end)
        # Printed and flushed BEFORE the request (not just after), so a
        # hung request is visible in the logs instead of leaving a long
        # silent gap that looks identical to "not started yet".
        print(f"  fetching {chart_type} {chunk_start} to {chunk_end} ...", file=sys.stderr, flush=True)
        chunk_rows = fetch_chunk(chart_type, areas, chunk_start, chunk_end)
        rows.extend(chunk_rows)
        print(f"  {chart_type} {chunk_start} to {chunk_end}: {len(chunk_rows)} rows", file=sys.stderr, flush=True)
        chunk_start = chunk_end + timedelta(days=1)
        chunk_count += 1
        if chunk_count % CHECKPOINT_EVERY_N_CHUNKS == 0 and chunk_start <= end:
            existing = merge_columns(existing, rows_to_wide(rows))
            rows = []
            save_callback(existing)
            print(f"  checkpoint saved ({len(existing)} rows so far)", file=sys.stderr, flush=True)
    existing = merge_columns(existing, rows_to_wide(rows))
    save_callback(existing)
    return existing


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    df.index.name = "datetime"
    return df


NOTES_LINES = [
    "UNITS",
    "Demand_Actual_MW / Demand_Forecast_MW: all-island (Republic of Ireland + Northern Ireland "
    "combined) system demand, instantaneous MW, 15-minute intervals.",
    "Wind_Actual_MW / Wind_Forecast_MW: all-island wind generation, instantaneous MW, 15-minute "
    "intervals.",
    "CO2_Intensity_gCO2_per_kWh: grams of CO2 per kWh of generation, all-island, 15-minute intervals "
    "(EirGrid/SONI's own published CO2 intensity - CO2_INTENSITY is often not reported for the most "
    "recent few hours of 'today', expect trailing NaNs there until it's published).",
    "",
    "SCOPE",
    "region=ALL (all-island, ROI + NI combined) only - the dashboard also supports ROI/NI "
    "individually, not pulled here.",
    "",
    "COVERAGE",
    f"Confirmed live back to {DATA_START.isoformat()} (not yet tested further back). Backfills in "
    f"{CHUNK_DAYS}-day chunks per chart type (confirmed the API accepts an arbitrary dateFrom/dateTo "
    "span with dateRange=day, unlike week/month/year which give inconsistent or empty results) - far "
    "fewer requests than one-per-day.",
    "",
    "SOURCE",
    f"EirGrid/SONI Smart Grid Dashboard's own chart API: {URL} - found via Playwright network "
    "capture, no login or key needed.",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    existing = load_archive(args.out)
    stop_at = date.today() - timedelta(days=1)  # today itself is still in progress
    os.makedirs(os.path.dirname(args.out), exist_ok=True)

    def save(df):
        xlsx_notes.write_workbook(args.out, {"Data": df}, NOTES_LINES, NOTES_SECTION_TITLES)

    # Each chart type is resumed independently (from the latest date that
    # already has non-null data for THAT type's own columns), not from a
    # single shared "last row" - otherwise, once one chart type's backfill
    # reaches near-today, the other two (not yet backfilled) would wrongly
    # be resumed from near-today too and their older history would never
    # get fetched.
    for chart_type, areas in CHART_TYPES.items():
        rd = resume_date_for(existing, chart_type)
        start = max(args.start_date, rd) if rd is not None else args.start_date
        if start > stop_at:
            print(f"{chart_type}: already current ({start} > {stop_at}) - skipping.", file=sys.stderr)
            continue
        print(f"{chart_type}: fetching {start} to {stop_at} ...", file=sys.stderr)
        existing = fetch_and_merge_chart_type(chart_type, areas, start, stop_at, existing, save)

    print(f"Saved to {args.out} ({len(existing)} rows, "
          f"{existing.index.min() if not existing.empty else 'n/a'} to "
          f"{existing.index.max() if not existing.empty else 'n/a'})")


if __name__ == "__main__":
    main()
