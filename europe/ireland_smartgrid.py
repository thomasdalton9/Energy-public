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
TIMEOUT = (10, 60)

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
    r = requests.get(URL, headers=HEADERS, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json().get("Rows", [])


def fetch_range(start, end):
    """One combined long DataFrame (datetime, FieldName, Value) across
    all three chart types, for [start, end] inclusive."""
    all_rows = []
    for chart_type, areas in CHART_TYPES.items():
        chunk_start = start
        while chunk_start <= end:
            chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS - 1), end)
            rows = fetch_chunk(chart_type, areas, chunk_start, chunk_end)
            all_rows.extend(rows)
            print(f"  {chart_type} {chunk_start} to {chunk_end}: {len(rows)} rows", file=sys.stderr)
            chunk_start = chunk_end + timedelta(days=1)
    if not all_rows:
        return pd.DataFrame()
    df = pd.DataFrame(all_rows)
    df["datetime"] = pd.to_datetime(df["EffectiveTime"], format="%d-%b-%Y %H:%M:%S")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    df["column"] = df["FieldName"].map(lambda f: FIELD_RENAME.get(f, f))
    wide = df.pivot_table(index="datetime", columns="column", values="Value", aggfunc="last")
    wide.index.name = "datetime"
    ordered = [c for c in FIELD_RENAME.values() if c in wide.columns]
    other = [c for c in wide.columns if c not in ordered]
    return wide[ordered + other].sort_index()


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    df.index.name = "datetime"
    return df


def upsert(existing, new_df):
    if new_df.empty:
        return existing
    if existing.empty:
        combined = new_df
    else:
        combined = pd.concat([existing, new_df])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "datetime"
    return combined.sort_index()


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
    if not existing.empty:
        resume_from = existing.index.max().date()  # re-pull the last saved day too (it may have been partial)
    else:
        resume_from = args.start_date
    stop_at = date.today() - timedelta(days=1)  # today itself is still in progress

    if resume_from > stop_at:
        print(f"Archive already current ({resume_from} > {stop_at}) - nothing to do.", file=sys.stderr)
        return

    print(f"Fetching {resume_from} to {stop_at} ...", file=sys.stderr)
    new_df = fetch_range(resume_from, stop_at)
    combined = upsert(existing, new_df)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved to {args.out} ({len(combined)} rows, {combined.index.min()} to {combined.index.max()})")


if __name__ == "__main__":
    main()
