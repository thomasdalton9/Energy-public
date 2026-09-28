"""
Pull South Africa's power generation mix by fuel type (coal, gas, oil,
nuclear, hydro, pumped storage, wind, solar, biomass) from Eskom's
public "Station Build Up" CSV, and maintain a growing local daily
archive.

Data source: https://www.eskom.co.za/dataportal/wp-content/uploads/{year}/{month}/Station_Build_Up.csv
This is NOT a historical archive to backfill from. Despite the URL
being dated by year/month, the file itself is always just a rolling
window of the last ~7 days of hourly readings - the source page is
literally titled "Station Build Up for the last 7 days", and Eskom's
own reference parser (in electricitymaps-contrib) refuses to look
back further than a week. There's no --from-date/--to-date here
because there's no range to request; Eskom simply doesn't expose
deeper history publicly.

Each hourly row already reports that hour's own generation value (not
a cumulative running total), so no subtraction is needed to isolate
"new" data - this script just re-fetches the current window on every
run and upserts by date into a local archive CSV (--out), keeping
whatever days it already has and adding/refreshing whatever the
current window shows. Run it at least once a week (the window is 7
days) so no day is ever skipped; running it more often (e.g. daily) is
harmless, just redundant. History only starts accumulating from
whenever you start running this script.
"""

import argparse
import csv
import io
import sys
from datetime import date, datetime, timezone

import openpyxl
import pandas as pd
import requests

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes/api_keys

import xlsx_notes

# No-cache headers plus a timestamp query param (see fetch_current_window):
# two scheduled runs three days apart got byte-identical windows ending
# 23 Sep, while Electricity Maps' own Eskom parser still uses this same
# URL - so a cached copy in front of Eskom's site is a plausible cause.
HEADERS = {
    "User-Agent": "gas-demand-scripts/1.0",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}

# Column index (within the row, after the leading datetime column) to
# category. Indices not listed here are storage/pumping-consumption/
# interruption columns, not generation, and are skipped.
COLUMN_CATEGORY = {
    0: "coal",  # Thermal_Gen_Excl_Pumping_and_SCO
    6: "nuclear",  # Nuclear_Generation
    8: "oil",  # Eskom_OCGT_Generation
    9: "gas",  # Eskom_Gas_Generation
    10: "oil",  # Dispatchable_IPP_OCGT
    11: "hydro",  # Hydro_Water_Generation
    12: "pumped_storage",  # Pumped_Water_Generation
    16: "wind",  # Wind
    17: "solar",  # PV
    18: "solar",  # CSP
    19: "biomass",  # Other_RE
}
CATEGORIES = ["coal", "gas", "oil", "nuclear", "hydro", "pumped_storage", "wind", "solar", "biomass"]


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def get_url():
    today = datetime.now(timezone.utc)
    return f"https://www.eskom.co.za/dataportal/wp-content/uploads/{today:%Y}/{today:%m}/Station_Build_Up.csv"


def fetch_current_window(session):
    """Fetch and parse Eskom's current rolling ~7-day hourly CSV."""
    response = session.get(get_url(), params={"t": int(datetime.now(timezone.utc).timestamp())}, timeout=30)
    response.raise_for_status()
    reader = csv.reader(io.StringIO(response.text))
    header = next(reader, None)
    if header is None:
        return pd.DataFrame()

    rows = []
    for row in reader:
        if not row or not row[0]:
            continue
        try:
            dt = pd.to_datetime(row[0])
        except (ValueError, TypeError):
            continue
        values = row[1:]
        entry = {"datetime": dt}
        for category in CATEGORIES:
            entry[category] = 0.0
        for index, category in COLUMN_CATEGORY.items():
            if index >= len(values):
                continue
            raw = (values[index] or "").strip()
            if not raw:
                continue
            try:
                entry[category] += max(float(raw), 0.0)
            except ValueError:
                continue
        rows.append(entry)

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame.from_records(rows).set_index("datetime").sort_index()


# Eskom's CSV includes rows for hours that haven't been reported yet
# (blank values), which the parser turns into all-zero hours. A day made
# mostly of those averaged down to near-zero and got saved as if it were
# real - so only hours with any generation count, and only days with a
# near-complete set of hours are kept.
MIN_HOURS_PER_DAY = 22


def to_daily_mean(hourly_df):
    if hourly_df.empty:
        return pd.DataFrame()
    reported = hourly_df[hourly_df[CATEGORIES].sum(axis=1) > 0]
    grouped = reported.groupby(reported.index.date)[CATEGORIES]
    counts = grouped.size()
    daily = grouped.mean()[counts >= MIN_HOURS_PER_DAY]
    daily.index.name = "date"
    return daily


def load_archive(path):
    # Reads the date from the first column whatever its header says - an
    # earlier version saved it without a header ("Unnamed: 0"), and
    # looking it up by name then failed, so the whole archive was
    # silently treated as empty and overwritten.
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except FileNotFoundError:
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df[df[CATEGORIES].sum(axis=1) > 0]


def upsert(existing, new):
    if existing.empty:
        return new
    if new.empty:
        return existing
    combined = pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


# Eskom's CSV normally refreshes daily. If the newest complete day it
# returns is older than this (Eskom itself runs ~4 days behind, so 6 leaves margin while still catching a real outage before its 7-day window rolls past), the source has gone stale (moved URL, stopped
# publishing) - fail the run so GitHub flags it, rather than quietly
# re-saving the same week forever.
STALE_AFTER_DAYS = 6


NOTES_LINES = [
    "UNITS",
    "Every category column is in MW, the daily mean of Eskom's hourly generation readings - "
    "'average MW for that day', not total daily energy.",
    "",
    "CATEGORIES",
    "coal, gas, oil, nuclear, hydro, wind, solar, biomass: generation by fuel/technology type.",
    "pumped_storage: pumped-hydro storage discharge.",
    "",
    "COVERAGE",
    "This is a rolling archive, not a one-shot historical pull - Eskom's source CSV is always "
    "just the last ~7 days, so this script upserts by date into the Data tab on every run, "
    "building up history from whenever you first started running it. Run it at least weekly "
    "so no day is skipped.",
    "",
    "SOURCE",
    "Eskom's public \"Station Build Up\" CSV.",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "COVERAGE", "SOURCE"}


def format_date_column(path):
    # pandas' default "YYYY-MM-DD" format in a narrow column was shown by
    # some viewers (e.g. phone previews) as "26/9" - year and month only -
    # making every September row look identical. A spelled-out month is
    # unambiguous in any viewer or locale.
    wb = openpyxl.load_workbook(path)
    ws = wb["Data"]
    for (cell,) in ws.iter_rows(min_row=2, max_col=1):
        cell.number_format = "dd-mmm-yyyy"
    ws.column_dimensions["A"].width = 14
    wb.save(path)


DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "south_africa_generation_mix_daily.xlsx")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    session = make_session()
    print(f"Fetching current Station Build Up window from {get_url()}...", file=sys.stderr)
    hourly = fetch_current_window(session)
    if hourly.empty:
        print("No data returned.", file=sys.stderr)
        sys.exit(1)
    new_daily = to_daily_mean(hourly)

    existing = load_archive(args.out)
    before_days = set(existing.index) if not existing.empty else set()
    combined = upsert(existing, new_daily)
    new_or_updated = sorted(set(combined.index) - before_days)

    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_column(args.out)
    print(f"Archive now has {len(combined)} days ({len(new_or_updated)} new since last run). Saved to {args.out}")
    print(combined.tail())

    latest = max(new_daily.index) if not new_daily.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(
            f"STALE SOURCE: newest complete day from Eskom is {latest} "
            f"({age} days old) - {get_url()} may have stopped updating or moved.",
            file=sys.stderr,
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
