"""
Fills the one gap MISO_FUEL_MIX_HISTORICAL_BACKFILL.py couldn't:
2026-01-01 to yesterday. That script used MISO's once-a-year
historical_gen_fuel_mix_<YEAR>.xlsx archive, which only appears after a
year closes (2025's was published 1/5/2026), so 2026's won't exist until
~Jan 2027.

MISO_HISTORICAL_FUEL_MIX_PLAYWRIGHT_DISCOVERY.py round 2 found a second,
DAILY report instead: <YYYYMMDD>_sr_gfm.xlsx at the same public
docs.misoenergy.org/marketreports/ path (no key needed), going back
~1368 days (per the report portal's own count) - roughly to 2023-01-01,
the same start date as the annual archive, suggesting this daily feed is
what the annual archive gets assembled from at year-end.

Confirmed via MISO_SR_GFM_INSPECT.py: a file named <YYYYMMDD>_sr_gfm.xlsx
has "Publish Date" = YYYYMMDD and "Market Date" = YYYYMMDD minus one day -
so to backfill market dates 2026-01-01..latest, this fetches filenames
2026-01-02..(latest+1), tolerating 404 for the very latest day(s) not
published yet.

Sheet "RT Generation Fuel Mix" (real-time actual - "DA Cleared Generation
Fuel Mix" is a day-ahead forecast, not used) has fixed-position columns
after 4 header rows (title, Publish Date, Market Date, region-group
labels) plus a 5th real header row: a per-region breakdown (Central,
North, South, each blank-separated) followed by a pre-summed *national*
total block - column 31 = HE (hour ending 1-24), 32-39 = Coal/Gas/
Nuclear/Hydro/Wind/Solar/Other/Storage, 40 = MISO (national Total MW).
Using that block directly means no manual cross-region summing (unlike
the annual archive's raw per-region rows) - confirmed identical column
layout across three widely-spaced sample dates (2026-01-01, 04-15, 09-29).

Fuel names renamed the same way as MISO_FUEL_MIX_HISTORICAL_BACKFILL.py
(Gas -> Natural Gas, Storage -> Battery Storage) so rows land in the same
archive columns; Hydro again has no live-feed equivalent and stays its
own column.

Reuses MISO_FUEL_MIX_DAILY.py's load_archive/upsert/format_date_column/
xlsx_notes so this merges into the exact same Data sheet - idempotent
(upsert by date), safe to re-run as more days publish.
"""

import argparse
import os
import sys
from datetime import date, timedelta

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib.util

_spec = importlib.util.spec_from_file_location(
    "miso_fuel_mix_daily", os.path.join(os.path.dirname(os.path.abspath(__file__)), "MISO_FUEL_MIX_DAILY.py")
)
miso_daily = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(miso_daily)

START_MARKET_DATE = date(2026, 1, 1)
MIN_HOURS_PER_DAY = miso_daily.MIN_HOURS_PER_DAY

DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(os.path.dirname(DIR), "output", "miso_fuel_mix_daily.xlsx")

# report download and parsing live in MISO_FUEL_MIX_DAILY.py (the daily pull reads the same report)
parse_day = miso_daily.parse_day


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--end", default=None, help="last market date to backfill (YYYY-MM-DD); default yesterday")
    args = parser.parse_args()

    end_day = date.fromisoformat(args.end) if args.end else date.today() - timedelta(days=1)

    print(f"Backfilling market dates {START_MARKET_DATE} to {end_day} from daily sr_gfm files ...",
          file=sys.stderr)

    rows = {}
    n_days = (end_day - START_MARKET_DATE).days + 1
    market_day = START_MARKET_DATE
    fetched, missing, incomplete = 0, 0, 0
    while market_day <= end_day:
        content = miso_daily.fetch_file(market_day)
        if content is None:
            missing += 1
            print(f"  {market_day}: {miso_daily.report_url(market_day)} not found (404) - skipping", file=sys.stderr)
        else:
            row, n_hours = parse_day(content, market_day)
            if row is None:
                incomplete += 1
                print(f"  {market_day}: only {n_hours}/24 hours - skipping", file=sys.stderr)
            else:
                rows[market_day] = row
                fetched += 1
        if (fetched + missing + incomplete) % 30 == 0:
            print(f"  ... {fetched + missing + incomplete}/{n_days} days processed "
                  f"({fetched} ok, {missing} missing, {incomplete} incomplete)", file=sys.stderr)
        market_day += timedelta(days=1)

    print(f"\nDone: {fetched} days fetched, {missing} missing (404), {incomplete} incomplete (<{MIN_HOURS_PER_DAY}h)",
          file=sys.stderr)

    if not rows:
        print("No days fetched - nothing to backfill.", file=sys.stderr)
        sys.exit(1)

    backfill_rows = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    backfill_rows.index.name = "date"

    existing = miso_daily.load_archive(args.out)
    before_days = set(existing.index) if not existing.empty else set()

    if existing.empty:
        combined = backfill_rows
    else:
        combined = pd.concat([existing, backfill_rows])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined = combined.sort_index()

    new_days = set(backfill_rows.index) - before_days
    overlap_days = set(backfill_rows.index) & before_days
    print(f"Added {len(new_days)} new day(s), {len(overlap_days)} day(s) already present "
          f"(kept existing/live values for those).")

    combined = miso_daily.reorder_columns(combined)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    miso_daily.xlsx_notes.write_workbook(
        args.out, {"Data": combined}, miso_daily.NOTES_LINES, miso_daily.NOTES_SECTION_TITLES
    )
    miso_daily.format_date_column(args.out)

    print(f"\nArchive now has {len(combined)} days total ({combined.index.min()} to "
          f"{combined.index.max()}). Saved to {args.out}")
    print(combined.tail())


if __name__ == "__main__":
    main()
