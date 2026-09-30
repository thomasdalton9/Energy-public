"""
One-time historical backfill for miso_fuel_mix_daily.xlsx (the archive
MISO_FUEL_MIX_DAILY.py maintains going forward from live /FuelMix/Today
pulls). MISO's real-time API only exposes Today/Yesterday - no historical
range query (confirmed by probing it) - but MISO separately publishes a
real, public, no-key annual archive at
https://docs.misoenergy.org/marketreports/historical_gen_fuel_mix_<YEAR>.xlsx,
found via MISO_HISTORICAL_FUEL_MIX_PLAYWRIGHT_DISCOVERY.py (the report
portal's grid is client-rendered, so the actual file names weren't visible
in the static HTML). Only 2023, 2024 and 2025 exist (2014-2022 all 404) -
confirmed by MISO_HISTORICAL_FUEL_MIX_INSPECT.py.

Each year's file has sheet "Gen Fuel Mix" with a few junk header/title/
legal-disclaimer rows before the real header (row 5 in the file, i.e.
pandas header=4), then one row per (Market Date, HourEnding 1-24, Region,
Fuel Type) with both a day-ahead ("DA Cleared UDS Generation") and a
real-time ("...RT Generation State Estimator", exact text has a stray
leading bracket in the source file) generation column, in MW. This uses
the RT column - the actual/measured generation, comparable to the
real-time API's ACT values the daily script uses (DA is a forecast/
schedule, not what happened).

This is hourly and by-region, not 5-minute and national like the live
feed - so to merge into the same archive: sum RT MW across all regions
for each (date, hour, fuel type) to get national hourly generation per
fuel type, then average the (up to) 24 hours to get a daily mean MW per
fuel type - the same "average MW that day" unit MISO_FUEL_MIX_DAILY.py
already uses. A day needs most of its 24 hours present to be trusted,
mirroring that script's 270/288 completeness threshold.

Fuel Type names in this archive differ slightly from the live feed's
CATEGORY names (this archive: Other, Coal, Gas, Wind, Nuclear, Hydro,
Solar, Storage - live feed: Coal, Natural Gas, Nuclear, Wind, Solar,
Battery Storage, Other, Imports). Renamed where the mapping is
unambiguous (Gas -> Natural Gas, Storage -> Battery Storage) so backfilled
and live rows land in the same columns; Hydro has no live-feed
equivalent (likely folded into the live feed's "Other") and Imports has
no historical equivalent (this archive is generation only, not net
interchange) - both are kept as their own columns and simply go NaN on
the side that doesn't have them, consistent with how the archive already
handles "whatever categories are actually present becomes a column".

Reuses MISO_FUEL_MIX_DAILY.py's own load_archive/upsert/format_date_column/
xlsx_notes plumbing so the backfilled rows merge into the exact same file
and sheet the daily automation writes to - drop_duplicates(keep="last")
means if a date already has a live-pulled row, the backfill won't
overwrite it.

Run once (workflow_dispatch only, no schedule). Safe to re-run: it's
idempotent (upsert on date), so re-running after MISO publishes a new
year's file just adds/refreshes that year.
"""

import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib.util

_spec = importlib.util.spec_from_file_location(
    "miso_fuel_mix_daily", os.path.join(os.path.dirname(os.path.abspath(__file__)), "MISO_FUEL_MIX_DAILY.py")
)
miso_daily = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(miso_daily)

BASE_URL = "https://docs.misoenergy.org/marketreports/historical_gen_fuel_mix_{year}.xlsx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)

YEARS = [2023, 2024, 2025]
EXPECTED_HOURS_PER_DAY = 24
MIN_HOURS_PER_DAY = 20  # mirrors MISO_FUEL_MIX_DAILY.py's 270/288 completeness bar

FUEL_TYPE_RENAME = {
    "Gas": "Natural Gas",
    "Storage": "Battery Storage",
}

DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(DIR, "miso_fuel_mix_daily.xlsx")


def fetch_year(year):
    url = BASE_URL.format(year=year)
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.content


def parse_year(content, year):
    tmp_path = os.path.join(DIR, f"_tmp_backfill_{year}.xlsx")
    with open(tmp_path, "wb") as f:
        f.write(content)
    try:
        df = pd.read_excel(tmp_path, sheet_name="Gen Fuel Mix", header=4)
    finally:
        os.remove(tmp_path)

    rt_col = next(c for c in df.columns if "RT Generation State Estimator" in str(c))
    df = df.rename(columns={rt_col: "RT_MW"})
    df = df[["Market Date", "HourEnding", "Region", "Fuel Type", "RT_MW"]].dropna(subset=["Market Date"])
    df["Market Date"] = pd.to_datetime(df["Market Date"]).dt.date
    df["Fuel Type"] = df["Fuel Type"].replace(FUEL_TYPE_RENAME)
    df["RT_MW"] = pd.to_numeric(df["RT_MW"], errors="coerce")
    return df


def to_daily_rows(df):
    """(date, hour, fuel type) national totals -> one row per date, same
    shape as MISO_FUEL_MIX_DAILY.to_daily_row()'s output."""
    national_hourly = (
        df.groupby(["Market Date", "HourEnding", "Fuel Type"])["RT_MW"].sum().reset_index()
    )
    hours_per_day = national_hourly.groupby("Market Date")["HourEnding"].nunique()

    daily_mean = (
        national_hourly.groupby(["Market Date", "Fuel Type"])["RT_MW"].mean().unstack("Fuel Type")
    )

    rows = {}
    skipped = []
    for day, n_hours in hours_per_day.items():
        if n_hours < MIN_HOURS_PER_DAY:
            skipped.append((day, n_hours))
            continue
        means = daily_mean.loc[day]
        row = {f"{cat}_MW": means[cat] for cat in means.index if pd.notna(means[cat])}
        total = sum(row.values())
        row["Total_MW"] = total
        renewable = row.get("Wind_MW", 0.0) + row.get("Solar_MW", 0.0)
        row["Renewables_Share"] = renewable / total if total else None
        row["Intervals_Reported"] = int(n_hours)  # hourly source: up to 24, not 288
        rows[day] = row

    if skipped:
        print(f"  skipped {len(skipped)} incomplete day(s) (fewer than {MIN_HOURS_PER_DAY}/"
              f"{EXPECTED_HOURS_PER_DAY} hours): {skipped[:5]}{'...' if len(skipped) > 5 else ''}",
              file=sys.stderr)

    out = pd.DataFrame.from_dict(rows, orient="index")
    out.index.name = "date"
    return out.sort_index()


def main():
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    all_years = []
    for year in YEARS:
        print(f"Fetching {year} ...", file=sys.stderr)
        content = fetch_year(year)
        if content is None:
            print(f"  {year}: not available (404) - skipping", file=sys.stderr)
            continue
        df = parse_year(content, year)
        print(f"  {year}: {len(df):,} raw rows, "
              f"{df['Market Date'].min()} to {df['Market Date'].max()}", file=sys.stderr)
        all_years.append(df)

    if not all_years:
        print("No historical years available - nothing to backfill.", file=sys.stderr)
        sys.exit(1)

    combined_raw = pd.concat(all_years, ignore_index=True)
    backfill_rows = to_daily_rows(combined_raw)
    print(f"\nBackfill produced {len(backfill_rows)} complete days "
          f"({backfill_rows.index.min()} to {backfill_rows.index.max()})")

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
    miso_daily.xlsx_notes.write_workbook(
        args.out, {"Data": combined}, miso_daily.NOTES_LINES, miso_daily.NOTES_SECTION_TITLES
    )
    miso_daily.format_date_column(args.out)

    print(f"\nArchive now has {len(combined)} days total ({combined.index.min()} to "
          f"{combined.index.max()}). Saved to {args.out}")
    print(combined.head())
    print("...")
    print(combined.tail())


if __name__ == "__main__":
    main()
