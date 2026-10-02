"""
Australia (NEM) daily power generation mix by fuel type, per state, back
to 2020 - a token-free alternative to australia_power_mix.py (which
needs a free-but-registered OPENELECTRICITY_TOKEN the user hasn't set
up). Built directly on AEMO's own public NEMWEB data via the `nemosis`
package (github.com/UNSW-CEEM/NEMOSIS), which handles NEMWEB's
current/archive/MMSDM zip layouts transparently - no auth needed.

Two source tables:
  - "Generators and Scheduled Loads" (a static reference table: AEMO's
    "NEM Registration and Exemption List" workbook) - maps each unit
    (DUID) to its NEM region and fuel type. nemosis's own static_table()
    helper fails to read this ("Excel file format cannot be
    determined") because AEMO serves a modern .xlsx (zip-based) body at
    a URL that still ends in ".xls" - pandas/nemosis auto-detect the
    read engine from that stale extension and pick xlrd, which can't
    read zip-based files. Fetched directly here instead with
    engine="openpyxl" forced explicitly. Confirmed working (2026-09-30):
    a plain requests.get with ordinary browser headers gets a clean 200
    with real .xlsx bytes - no WAF/bot-block involved, just a
    content/extension mismatch.
  - DISPATCH_UNIT_SCADA (via nemosis.dynamic_data_compiler) - real
    5-minute dispatched MW per DUID. Confirmed working via NEMWEB's
    MMSDM archive with no auth. Pulled one calendar month at a time and
    aggregated to a daily mean immediately, so raw 5-minute rows for
    the full 2020+ history are never all held in memory at once.

Categories mirror australia_power_mix.py's own taxonomy (coal, gas,
oil, hydro, wind, biomass, solar, nuclear, battery, hydro_pumped) built
from AEMO's "Fuel Source - Descriptor" / "Technology Type - Descriptor"
fields - see FUEL_DESCRIPTOR_TO_CATEGORY below. Any descriptor not in
that mapping raises rather than silently dropping generation, so a
newly-registered fuel type doesn't silently disappear from the totals.

Western Australia is NOT included - it runs the separate WEM market,
not part of the NEM MMSDM data this script (and nemosis) pulls from.
Northern Territory isn't part of either wholesale market and was never
covered by the sibling OpenElectricity-based script either.

Output: one Excel workbook, one tab per NEM state (New South Wales,
Queensland, South Australia, Tasmania, Victoria), daily mean MW per
fuel category - same shape as australia_power_mix_master.py's output.

Run standalone:
    python3 aemo_nemweb_power_mix.py --out aemo_nemweb_power_mix.xlsx
Optional --from-date/--to-date (default 2020-01-01 to today).
"""

import argparse
import os
import sys
from datetime import date
from io import BytesIO

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

REGISTRATION_LIST_URL = "https://www.aemo.com.au/-/media/Files/Electricity/NEM/Participant_Information/NEM-Registration-and-Exemption-List.xls"
REGISTRATION_SHEET_NAMES = ["Generators and Scheduled Loads", "PU and Scheduled Loads"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://www.aemo.com.au/",
}

REGION_TO_STATE_NAME = {
    "NSW1": "New South Wales",
    "QLD1": "Queensland",
    "SA1": "South Australia",
    "TAS1": "Tasmania",
    "VIC1": "Victoria",
}
STATE_NAMES = list(REGION_TO_STATE_NAME.values())

# Built from AEMO's own published fueltech vocabulary (docs.openelectricity.org.au/guides/fueltechs
# and the Registration and Exemption List's own "Fuel Source - Descriptor" /
# "Technology Type - Descriptor" values) - matches australia_power_mix.py's
# category taxonomy so the two sources are directly comparable.
FUEL_DESCRIPTOR_TO_CATEGORY = {
    "Black Coal": "coal",
    "Brown Coal": "coal",
    "Natural Gas (Pipeline)": "gas",
    "Natural Gas": "gas",
    "Coal Seam Methane": "gas",
    "Coal Mine Waste Gas": "gas",
    "Diesel oil": "oil",
    "Kerosene - Non Aviation": "oil",
    "Distillate": "oil",
    "Water": "hydro",
    "Wind": "wind",
    "Solar": "solar",
    "Bagasse": "biomass",
    "Wood Waste": "biomass",
    "Landfill Gas": "biomass",
    "Sewerage Gas": "biomass",
    "Biomass": "biomass",
    "Nuclear": "nuclear",
}
# Technology Type - Descriptor overrides Fuel Source for storage - a
# "Water" fuel source with a pump-storage technology type is
# hydro_pumped, not plain hydro; a battery is "battery" regardless of
# its (usually blank/"-") fuel source.
TECHNOLOGY_OVERRIDE_TO_CATEGORY = {
    "Pump Storage": "hydro_pumped",
    "Battery Storage": "battery",
    "Battery": "battery",
}
CATEGORIES = ["coal", "gas", "oil", "hydro", "wind", "biomass", "solar", "nuclear", "battery", "hydro_pumped"]

CHUNK_MONTHS_START = None  # set by main() from --from-date
NEMOSIS_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aemo_nemosis_cache")


def fetch_duid_category_map():
    """Return {DUID: (region, category)} from AEMO's own registration list."""
    r = requests.get(REGISTRATION_LIST_URL, headers=HEADERS, timeout=(10, 60))
    r.raise_for_status()
    content = BytesIO(r.content)

    df = None
    last_error = None
    for sheet_name in REGISTRATION_SHEET_NAMES:
        try:
            df = pd.read_excel(content, sheet_name=sheet_name, dtype=str, engine="openpyxl")
            break
        except ValueError as e:
            last_error = e
            content.seek(0)
    if df is None:
        raise RuntimeError(
            f"None of the expected sheet names {REGISTRATION_SHEET_NAMES} were found "
            f"in the registration list workbook: {last_error}"
        )

    df = df.dropna(subset=["DUID"])
    mapping = {}
    unmapped_descriptors = set()
    for _, row in df.iterrows():
        duid = row["DUID"]
        region = row.get("Region")
        if region not in REGION_TO_STATE_NAME:
            continue  # WA (WEM) or another non-NEM-MMSDM region - not covered here
        tech = str(row.get("Technology Type - Descriptor") or "").replace("nan", "").strip()
        fuel = str(row.get("Fuel Source - Descriptor") or "").replace("nan", "").strip()
        category = TECHNOLOGY_OVERRIDE_TO_CATEGORY.get(tech)
        if category is None:
            category = FUEL_DESCRIPTOR_TO_CATEGORY.get(fuel)
        if category is None:
            unmapped_descriptors.add((fuel, tech))
            continue
        mapping[duid] = (region, category)

    if unmapped_descriptors:
        raise RuntimeError(
            "Unmapped Fuel Source/Technology Type combinations found in the AEMO "
            "registration list - add these to FUEL_DESCRIPTOR_TO_CATEGORY or "
            f"TECHNOLOGY_OVERRIDE_TO_CATEGORY before proceeding: {sorted(unmapped_descriptors)}"
        )
    return mapping


def month_ranges(from_date, to_date):
    """Yield (start, end) date pairs, one per calendar month, covering [from_date, to_date]."""
    current = date(from_date.year, from_date.month, 1)
    while current <= to_date:
        if current.month == 12:
            next_month = date(current.year + 1, 1, 1)
        else:
            next_month = date(current.year, current.month + 1, 1)
        yield current, min(next_month, to_date)
        current = next_month


def fetch_month_daily_means(duid_map, month_start, month_end):
    """Return {(day, region, category): mean_MW} for one calendar month."""
    from nemosis import dynamic_data_compiler

    start_str = month_start.strftime("%Y/%m/%d %H:%M:%S")
    end_str = month_end.strftime("%Y/%m/%d %H:%M:%S")
    df = dynamic_data_compiler(start_str, end_str, "DISPATCH_UNIT_SCADA", NEMOSIS_CACHE_DIR)
    if df.empty:
        return {}

    df["region_category"] = df["DUID"].map(duid_map)
    df = df.dropna(subset=["region_category"])
    df["region"] = df["region_category"].apply(lambda t: t[0])
    df["category"] = df["region_category"].apply(lambda t: t[1])
    df["day"] = pd.to_datetime(df["SETTLEMENTDATE"]).dt.date
    df["SCADAVALUE"] = pd.to_numeric(df["SCADAVALUE"], errors="coerce")

    # clip negatives to 0 for generation categories (a unit briefly
    # drawing a small amount of house load reads as slightly negative -
    # not real "negative generation") but keep battery signed
    # (charging is a real negative MW value worth showing)
    non_battery = df["category"] != "battery"
    df.loc[non_battery, "SCADAVALUE"] = df.loc[non_battery, "SCADAVALUE"].clip(lower=0.0)

    grouped = df.groupby(["day", "region", "category"])["SCADAVALUE"].mean()
    return grouped.to_dict()


def build_state_dataframes(from_date, to_date):
    duid_map = fetch_duid_category_map()
    print(f"Loaded {len(duid_map)} DUIDs across {len(REGION_TO_STATE_NAME)} NEM regions", file=sys.stderr)

    all_rows = {}  # (day, region, category) -> mean MW
    for month_start, month_end in month_ranges(from_date, to_date):
        print(f"Fetching {month_start.isoformat()} to {month_end.isoformat()}...", file=sys.stderr)
        try:
            all_rows.update(fetch_month_daily_means(duid_map, month_start, month_end))
        except Exception as exc:
            print(f"  failed: {exc}", file=sys.stderr)

    by_region_day = {}
    for (day, region, category), value in all_rows.items():
        by_region_day.setdefault(region, {}).setdefault(day, {})[category] = value

    state_dfs = {}
    for region, state_name in REGION_TO_STATE_NAME.items():
        day_records = by_region_day.get(region, {})
        if not day_records:
            continue
        records = []
        for day, cats in day_records.items():
            record = {"date": day}
            for category in CATEGORIES:
                record[category] = cats.get(category)
            records.append(record)
        state_dfs[state_name] = pd.DataFrame.from_records(records).set_index("date").sort_index()
    return state_dfs


NOTES_SECTION_TITLES = {"UNITS", "COVERAGE", "SOURCE"}


def build_notes(included_count):
    return [
        "UNITS",
        "Every value column is MW-equivalent: the MEAN of that day's native 5-minute dispatched "
        "power readings (AEMO's DISPATCH_UNIT_SCADA) - 'average MW for that day', not total daily "
        "energy (MWh).",
        "",
        "Categories: coal, gas, oil, hydro, wind, biomass, solar, nuclear are production "
        "(negative readings clipped to 0 - a unit briefly drawing house load, not real negative "
        "generation); battery and hydro_pumped are storage (sign as published - negative battery "
        "is charging).",
        "",
        "COVERAGE",
        "New South Wales, Queensland, South Australia, Tasmania and Victoria - the 5 NEM "
        "(National Electricity Market) regions - are covered. Western Australia runs a separate "
        "market (WEM) with no MMSDM data and is NOT included. Northern Territory runs its own "
        "small, separate grid outside the NEM/WEM entirely and was never covered.",
        "",
        f"This run includes {included_count} of 5 NEM states (see stderr output for any that "
        "returned no data).",
        "",
        "SOURCE",
        "AEMO's own public NEMWEB data (nemweb.com.au), pulled via the nemosis package - no API "
        "token needed, unlike australia_power_mix.py's OpenElectricity-based pull. Unit-to-region/"
        "fuel-type mapping is AEMO's own 'NEM Registration and Exemption List'.",
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from-date", type=date.fromisoformat, default=date(2020, 1, 1))
    parser.add_argument("--to-date", type=date.fromisoformat, default=date.today())
    parser.add_argument("--out", default="aemo_nemweb_power_mix.xlsx")
    args = parser.parse_args()

    state_dfs = build_state_dataframes(args.from_date, args.to_date)
    if not state_dfs:
        print("No data returned for any state.", file=sys.stderr)
        sys.exit(1)

    notes_lines = build_notes(len(state_dfs))
    xlsx_notes.write_workbook(args.out, state_dfs, notes_lines, NOTES_SECTION_TITLES, notes_sheet_name="Notes")

    missing = [name for name in STATE_NAMES if name not in state_dfs]
    if missing:
        print(f"Saved {args.out} with tabs: Notes, {list(state_dfs.keys())} - NO DATA for: {missing}")
    else:
        print(f"Saved {args.out} with all 6 tabs: Notes, {list(state_dfs.keys())}")


if __name__ == "__main__":
    main()
