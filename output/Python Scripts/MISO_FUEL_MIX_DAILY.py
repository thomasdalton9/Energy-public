"""
MISO generation by fuel type, daily archive (average MW per fuel per day).

Data source: MISO's daily real-time generation fuel mix report, public, no key:
  https://docs.misoenergy.org/marketreports/<YYYYMMDD>_sr_gfm.xlsx
A file named <YYYYMMDD> is published the next morning for market date YYYYMMDD minus one day (its
"Publish Date" / "Market Date" header rows). Sheet "RT Generation Fuel Mix" (real-time actuals; the
"DA Cleared" sheet is a day-ahead forecast and is not used) has 24 hourly rows (HE 1-24) and a pre-summed
national block: Coal, Gas, Nuclear, Hydro, Wind, Solar, Other, Storage, then the MISO total.

Incremental: each run fetches only market dates after the last one saved, up to yesterday (at most
MAX_DAYS_PER_RUN per run, so a long outage catches up over a few runs). The same report backs the
2023-2025 annual archive (MISO_FUEL_MIX_HISTORICAL_BACKFILL.py) and the 2026 backfill
(MISO_FUEL_MIX_2026_GAP_BACKFILL.py), so the whole series is one consistent source; it agrees with
EIA-930's MISO series to within about half a percent a day.

Until Oct 2026 this script read MISO's live API (public-api.misoenergy.org/api/FuelMix/Today) late each
evening instead. That endpoint has no Hydro, adds Imports to its total and covers only the day so far,
so its two saved days (29-30 Sep 2026) didn't match the rest of the series; they are replaced by the
report's figures.
"""

import argparse
import io
import os
import sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

BASE = "https://docs.misoenergy.org/marketreports/{name}"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)

FIRST_MARKET_DATE = date(2023, 1, 1)   # the report archive's start
MAX_DAYS_PER_RUN = 60
MIN_HOURS_PER_DAY = 20                 # a day with fewer hourly rows is skipped, not saved thin

# fixed column positions in "RT Generation Fuel Mix" (MISO_SR_GFM_INSPECT.py, confirmed on dates across 2026)
COL_HE = 31
COL_CATEGORIES = {
    "Coal": 32, "Natural Gas": 33, "Nuclear": 34, "Hydro": 35,
    "Wind": 36, "Solar": 37, "Other": 38, "Battery Storage": 39,
}
COL_TOTAL = 40

RENEWABLE_CATEGORIES = ["Wind", "Solar"]


def report_url(market_day):
    return BASE.format(name=f"{(market_day + timedelta(days=1)).strftime('%Y%m%d')}_sr_gfm.xlsx")


def fetch_file(market_day):
    """The report for one market date (bytes), or None if MISO hasn't published it (404)."""
    r = requests.get(report_url(market_day), headers=HEADERS, timeout=TIMEOUT)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.content


def parse_day(content, market_day):
    """Daily mean MW per fuel from the national block of the real-time sheet -> (row, hours found)."""
    raw = pd.read_excel(io.BytesIO(content), sheet_name="RT Generation Fuel Mix", header=None)
    he = pd.to_numeric(raw[COL_HE], errors="coerce")
    rows = raw[he.between(1, 24)]
    n_hours = len(rows)
    if n_hours < MIN_HOURS_PER_DAY:
        return None, n_hours
    means = {cat: pd.to_numeric(rows[col], errors="coerce").mean() for cat, col in COL_CATEGORIES.items()}
    total = pd.to_numeric(rows[COL_TOTAL], errors="coerce").mean()
    row = {f"{cat}_MW": v for cat, v in means.items() if pd.notna(v)}
    row["Total_MW"] = total if pd.notna(total) else sum(row.values())
    renewable = sum(means.get(c, 0.0) for c in RENEWABLE_CATEGORIES)
    row["Renewables_Share"] = renewable / row["Total_MW"] if row["Total_MW"] else None
    row["Intervals_Reported"] = n_hours
    return pd.Series(row, name=market_day), n_hours


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new_row):
    if new_row is None:
        return existing
    new_df = new_row.to_frame().T
    new_df.index.name = "date"
    if existing.empty:
        combined = new_df
    else:
        combined = pd.concat([existing, new_df])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


# Display order for the *_MW columns - not the order MISO's API reports
# categories in (that varies), just how the archive reads left to right:
# baseload/dispatchable first, then renewables, then storage, then the
# catch-all. Hydro and any category not in this list still show up (per
# the CATEGORIES note below) - they're just placed after the ones named
# here rather than wherever alphabetical/pivot order would put them.
CATEGORY_DISPLAY_ORDER = ["Nuclear", "Coal", "Natural Gas", "Hydro", "Wind", "Solar", "Battery Storage", "Other"]


def reorder_columns(df):
    named_cat_cols = [f"{c}_MW" for c in CATEGORY_DISPLAY_ORDER if f"{c}_MW" in df.columns]
    other_cat_cols = [c for c in df.columns if c.endswith("_MW") and c not in named_cat_cols]
    non_cat_cols = [c for c in df.columns if not c.endswith("_MW")]
    return df[named_cat_cols + other_cat_cols + non_cat_cols]


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
    "All *_MW columns are MW, the daily mean of MISO's 24 hourly real-time readings for that fuel - 'average "
    "MW for that day', not daily energy (x 24 for MWh). Total_MW is MISO's own national total (generation only, "
    "no imports).",
    "",
    "CATEGORIES",
    "Nuclear, Coal, Natural Gas, Hydro, Wind, Solar, Battery Storage (net: negative when charging), Other - "
    "the national block of MISO's 'RT Generation Fuel Mix' report.",
    "Renewables_Share: (Wind + Solar) daily mean, as a share of Total_MW.",
    "Intervals_Reported: hourly rows found for that day (24 for a complete day; a day with fewer than "
    f"{MIN_HOURS_PER_DAY} is not saved).",
    "",
    "TIMESTAMPS",
    "The 'date' index is MISO's market day (Eastern Standard Time, hour ending 1-24).",
    "",
    "COVERAGE",
    "Daily from 2023-01-01 (the report archive's start). Each report is published the morning after the "
    "market day; each run adds the days published since the last saved one.",
    "",
    "SOURCE",
    "MISO (Midcontinent Independent System Operator) daily real-time generation fuel mix report: "
    "https://docs.misoenergy.org/marketreports/<YYYYMMDD>_sr_gfm.xlsx (market reports, public, no key).",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "TIMESTAMPS", "COVERAGE", "SOURCE"}

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "miso_fuel_mix_daily.xlsx")

# The report appears the morning after each market day; flag it if the newest saved day falls this far behind
STALE_AFTER_DAYS = 4


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    existing = load_archive(args.out)
    if not existing.empty:
        # rows from the old live 5-minute pull (>24 'intervals') don't match the report series: refetch them
        live = existing.index[pd.to_numeric(existing.get("Intervals_Reported"), errors="coerce") > 24]
        if len(live):
            print(f"Replacing {len(live)} day(s) saved from the old live API: {sorted(live)}", file=sys.stderr)
            existing = existing.drop(index=live)
    yesterday = datetime.now(ZoneInfo("America/New_York")).date() - timedelta(days=1)
    start = max(existing.index) + timedelta(days=1) if not existing.empty else FIRST_MARKET_DATE
    days = [start + timedelta(days=i) for i in range((yesterday - start).days + 1)][:MAX_DAYS_PER_RUN]
    print(f"Fetching market dates {days[0] if days else '-'} to {days[-1] if days else '-'} "
          f"({len(days)} day(s))", file=sys.stderr)

    combined = existing
    for d in days:
        content = fetch_file(d)
        if content is None:
            print(f"  {d}: not published yet ({report_url(d)} 404)", file=sys.stderr)
            break   # later days can't be out either
        row, n_hours = parse_day(content, d)
        if row is None:
            print(f"  {d}: only {n_hours}/24 hours - skipped", file=sys.stderr)
            continue
        combined = upsert(combined, row)
        print(f"  {d}: {n_hours} hours, total {row['Total_MW']:,.0f} MW average")

    combined = reorder_columns(combined.dropna(axis=1, how="all"))   # e.g. Imports_MW, only the old live rows had it
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_column(args.out)
    print(f"Archive now has {len(combined)} days. Saved to {args.out}")
    print(combined.tail())

    latest = max(combined.index) if not combined.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest saved day is {latest} ({age} days old) - the sr_gfm report may have "
              "moved or stopped publishing.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
