"""
Pull MISO's real-time generation by fuel type and maintain a growing
daily archive.

Data source: https://public-api.misoenergy.org/api/FuelMix/Today
Found via MISO_REACHABILITY_DISCOVERY.py - genuinely open, no API key
or account needed (unlike ERCOT/PJM's gated APIs). Returns every
5-minute interval (INTERVALEST) so far for the current MISO operating
day, each interval reporting instantaneous MW ("ACT") per fuel
category (Coal, Natural Gas, Nuclear, Wind, Solar, Battery Storage,
Other, Imports).

Originally used /FuelMix/Yesterday instead, on the assumption its name
meant a complete prior day - but two live pulls showed its generation-
by-fuel categories (everything except Imports) actually mirror
/Today's still-in-progress current day, not a real "yesterday" (only
Imports was genuinely dated the day before). Rather than depend on
that inconsistency, this uses /Today directly and is scheduled to run
late in the Eastern day (see the workflow: two UTC times, one per DST
regime, so one of them always lands close to Eastern midnight)  so
the "day so far" is close enough to complete to pass the interval-
coverage threshold below.

Categories are read from whatever the API actually reports each run,
not hardcoded, so a new one MISO adds shows up automatically as a new
archive column. If a single payload ever mixes rows from more than
one calendar date again (as /Yesterday's did), this picks whichever
date the most distinct categories actually have data for - see
to_daily_row()'s docstring.

INTERVALEST timestamps are Eastern local time (MISO's own operating
day), kept as reported - not converted to UTC.

Like south_africa_generation_mix.py, this upserts by date into a
local archive on every run: MISO's API is not a historical range
query (no /FuelMix/{date} - confirmed by probing it), just "today"
and "yesterday", so history only starts accumulating from whenever
this script started running. Run it at least once a day, late in the
Eastern day, so no day is skipped or saved too early/incomplete.
"""

import argparse
import os
import sys
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://public-api.misoenergy.org/api/FuelMix/Today"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
}
TIMEOUT = (10, 30)

# MISO reports every 5 minutes -> 288 intervals in a complete day.
# A day is only accepted if most of that window actually came back;
# a day cut short (API hiccup, this script started mid-day) keeps
# whatever the archive already had instead of overwriting it with a
# thin, misleading mean.
EXPECTED_INTERVALS_PER_DAY = 288
MIN_INTERVALS_PER_DAY = 270

RENEWABLE_CATEGORIES = ["Wind", "Solar"]


def fetch_today():
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def parse_intervals(payload):
    """Long list of {INTERVALEST, CATEGORY, ACT, ...} -> wide DataFrame,
    index = interval datetime, columns = fuel category, values = MW."""
    rows = payload.get("Fuel", {}).get("Type", [])
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    # e.g. "2026-09-29 7:15:00 PM" - explicit format avoids pandas
    # falling back to per-row dateutil parsing (slow, and warns).
    df["INTERVALEST"] = pd.to_datetime(df["INTERVALEST"], format="%Y-%m-%d %I:%M:%S %p")
    df["ACT"] = pd.to_numeric(df["ACT"], errors="coerce")
    wide = df.pivot_table(index="INTERVALEST", columns="CATEGORY", values="ACT", aggfunc="last")
    return wide.sort_index()


def to_daily_row(wide, expected_day):
    """Mean MW per category for whichever date the most fuel
    categories actually agree on, plus Total_MW (sum of categories,
    matching the API's own TotalMW) and a renewables share. Only for
    days with enough intervals present.

    Two live pulls of "Yesterday" both mixed two different calendar
    dates into one payload, but not the way a naive "which date has
    more rows" or "trust the expected Eastern date" check assumes.
    One pull had Imports on 288 timestamps and the other 7 generation
    categories on a disjoint 237 timestamps; filtering to the
    Eastern-computed expected day picked Imports' date (it happened to
    be the real "yesterday") while every generation category turned
    out to be dated *today* instead - a genuine inconsistency in how
    MISO itself labels different categories in this endpoint, not a
    which-date-to-trust bug. Trusting either a single category's date
    or a precomputed wall-clock date silently drops whichever
    categories disagree.

    Instead, pick the date where the most distinct categories actually
    have data (ties broken by total row count) - empirically the day
    the payload is really "about", regardless of which category's
    labeling is off - then average only that date's rows (an
    individual category still missing entirely for that date ends up
    correctly NaN rather than a fabricated value)."""
    if wide.empty:
        return None, 0

    dates = wide.index.date
    coverage_by_date = wide.notna().groupby(dates).sum()  # date -> per-category non-null count
    categories_present = (coverage_by_date > 0).sum(axis=1)  # date -> how many distinct categories
    row_counts = pd.Series(dates).value_counts()

    chosen_day = sorted(
        categories_present.index,
        key=lambda d: (categories_present[d], row_counts.get(d, 0)),
    )[-1]

    if chosen_day != expected_day:
        print(f"  note: payload's best-covered date is {chosen_day}, not the expected "
              f"{expected_day} - using {chosen_day} (categories present: "
              f"{categories_present[chosen_day]}/{wide.shape[1]})", file=sys.stderr)
    other_days = sorted(set(dates) - {chosen_day})
    if other_days:
        print(f"  note: payload also contained data for other date(s), ignored: {other_days}",
              file=sys.stderr)

    on_chosen_day = wide[dates == chosen_day]
    n_intervals = len(on_chosen_day)
    if n_intervals < MIN_INTERVALS_PER_DAY:
        return None, n_intervals

    means = on_chosen_day.mean()
    row = {f"{cat}_MW": means[cat] for cat in means.index}
    total = means.sum()
    row["Total_MW"] = total
    renewable = sum(means.get(cat, 0.0) for cat in RENEWABLE_CATEGORIES)
    row["Renewables_Share"] = renewable / total if total else None
    row["Intervals_Reported"] = n_intervals
    return pd.Series(row, name=chosen_day), n_intervals


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
    "All *_MW columns are MW, the daily mean of MISO's 5-minute generation readings for that fuel "
    "category - 'average MW for that day', not total daily energy. Total_MW is the sum of the "
    "category means (matches the sum of the API's own instantaneous category values).",
    "",
    "CATEGORIES",
    "Whatever MISO's API itself reports (typically Coal, Natural Gas, Nuclear, Wind, Solar, "
    "Battery Storage, Other, Imports) - not a fixed list maintained here, so a category MISO adds "
    "shows up as a new column automatically. Columns are ordered Nuclear, Coal, Natural Gas, Hydro, "
    "Wind, Solar, Battery Storage, Other, then any other category, then Total_MW/Renewables_Share/"
    "Intervals_Reported.",
    "Renewables_Share: (Wind + Solar) daily mean, as a share of Total_MW.",
    "Intervals_Reported: how many of the expected 288 five-minute intervals that day actually came "
    "back - always close to 288 for a complete day; a day is only saved if at least "
    f"{MIN_INTERVALS_PER_DAY} came back.",
    "",
    "TIMESTAMPS",
    "The 'date' index is MISO's own operating day (Eastern local time, as the API reports it - not "
    "converted to UTC).",
    "",
    "COVERAGE",
    "This is NOT a historical range API - MISO's public-api.misoenergy.org only exposes 'Today' and "
    "'Yesterday' (and 'Yesterday' turned out to actually mirror 'Today' for every category except "
    "Imports, not a real prior day - see the script docstring), no query-by-date endpoint (confirmed "
    "by probing it), so this script upserts by date into the Data tab on every run, building up "
    "history from whenever it first started running. Scheduled late in the Eastern day so 'today so "
    "far' is close enough to complete to save.",
    "",
    "SOURCE",
    f"MISO (Midcontinent Independent System Operator) public real-time API: {URL}",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "TIMESTAMPS", "COVERAGE", "SOURCE"}

DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs", "miso_fuel_mix_daily.xlsx")

# MISO's own API is real-time; if this script's own run misses several
# days in a row (workflow disabled, MISO API down), flag it rather
# than silently going stale forever.
STALE_AFTER_DAYS = 4


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    target_day = datetime.now(ZoneInfo("America/New_York")).date()
    print(f"Fetching {URL} (expecting {target_day}) ...", file=sys.stderr)
    payload = fetch_today()
    wide = parse_intervals(payload)
    new_row, n_intervals = to_daily_row(wide, target_day)

    existing = load_archive(args.out)
    before_days = set(existing.index) if not existing.empty else set()
    combined = upsert(existing, new_row)
    is_new = bool(new_row is not None and new_row.name not in before_days)

    if new_row is None:
        print(f"Not enough intervals to save a day ({n_intervals}/{EXPECTED_INTERVALS_PER_DAY}) - "
              "keeping the archive as-is.", file=sys.stderr)
    else:
        print(f"Day {new_row.name}: {n_intervals}/{EXPECTED_INTERVALS_PER_DAY} intervals "
              f"({'new' if is_new else 'refreshed'})")

    combined = reorder_columns(combined)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_column(args.out)

    print(f"Archive now has {len(combined)} days. Saved to {args.out}")
    print(combined.tail())

    latest = max(combined.index) if not combined.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest saved day is {latest} ({age} days old) - "
              f"{URL} may have changed or stopped updating.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
