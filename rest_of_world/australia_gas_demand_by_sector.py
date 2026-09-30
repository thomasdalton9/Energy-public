"""
Pull Australia's (east coast + Northern Territory) daily gas demand by
end-use sector from AEMO's Gas Bulletin Board (GBB) full-history flow
archive, and save it to an Excel workbook with one tab per state (each
with the same 4 daily TJ sector columns), a "Total" tab (all states
summed) last, and a highlighted "Units" tab explaining what the
columns mean.

Note: the source archive is granular by facility (141 facilities x
~2,900 gas days) - this script aggregates that away into just 4 daily
sector totals per state; no per-facility rows are kept in the output.
State comes directly from the "State" field AEMO publishes on each
flow row, not from a separate lookup.

Data source: https://nemweb.com.au/Reports/Current/GBB/
  - GasBBActualFlowStorage.zip: full-history archive of daily flow/
    demand per facility per gas day (public, no API key), covering
    2018-09-29 onward.
  - GasBBFacilities.CSV: reference list mapping each facility to its
    type (pipeline, production, storage, LNG export, gas-powered
    generation, large user, blended distribution, compression).

Sector categories (end-use facility types only - summing every
facility type would double-count gas that pipelines also report as
"demand" simply because they're transporting it, not consuming it):
  - gas_powered_generation (AEMO type BBGPG)
  - large_industrial_users (BBLARGE)
  - blended_distribution (BDIST) - residential/commercial city-gas
  - lng_export (LNGEXPORT)

IMPORTANT: AEMO expanded the Bulletin Board's reporting scope on
2023-03-15. Before that date, NONE of these four end-use facility
types reported at all - it's a real reporting gap, not missing data.
Any chart of this series across that boundary shows demand rising
from zero, which is a reporting change, not a market event. Only
pipeline/production/storage flows are meaningful before 2023-03-15,
and those aren't fetched by this script (it's scoped to end-use
sectors only).

The "gas day" (AEMO's GasDate) runs 06:00-06:00 fixed Australian
Eastern Standard Time (UTC+10, no daylight saving) and is used
directly as published, not reconstructed from timestamps.

Units are terajoules (TJ), as published - not converted to MW/MWh
like the other scripts in this repo.
"""

import argparse
import io
import sys
import zipfile
from datetime import date

import pandas as pd
import requests

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes/api_keys

import xlsx_notes

GBB_BASE = "https://nemweb.com.au/Reports/Current/GBB/"
FLOW_ARCHIVE_URL = GBB_BASE + "GasBBActualFlowStorage.zip"
FACILITIES_URL = GBB_BASE + "GasBBFacilities.CSV"

HEADERS = {"User-Agent": "gas-demand-scripts/1.0 (+github.com/thomasdalton9/Energy)"}

FACILITY_TYPE_TO_SECTOR = {
    "BBGPG": "gas_powered_generation",
    "BBLARGE": "large_industrial_users",
    "BDIST": "blended_distribution",
    "LNGEXPORT": "lng_export",
}
SECTORS = list(FACILITY_TYPE_TO_SECTOR.values())

GBB_EXPANSION_DATE = date(2023, 3, 15)


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def fetch_facilities(session):
    """Return {facility_id: facility_type}."""
    response = session.get(FACILITIES_URL, timeout=60)
    response.raise_for_status()
    text = response.content.decode("utf-8-sig", errors="replace")
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip().lower() for c in df.columns]
    return dict(zip(df["facilityid"], df["facilitytype"].astype(str).str.strip()))


def fetch_flow_archive(session):
    """Download and unzip the full-history flow archive; return raw CSV text."""
    response = session.get(FLOW_ARCHIVE_URL, timeout=180)
    response.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
        members = [name for name in zf.namelist() if not name.endswith("/")]
        if len(members) != 1:
            raise RuntimeError(f"Expected one file in the flow archive, found {len(members)}: {members[:5]}")
        return zf.read(members[0]).decode("utf-8-sig", errors="replace")


def _pivot_daily_sector(df):
    """(date, sector) -> demand_tj, with every sector column present even if empty."""
    daily = df.groupby(["gas_date", "sector"])["demand_tj"].sum().unstack("sector")
    for sector in SECTORS:
        if sector not in daily.columns:
            daily[sector] = 0.0
    daily = daily[SECTORS].sort_index()
    daily.index.name = "date"
    return daily


def aggregate_daily_sector_demand_by_state(flow_text, facility_types):
    """Sum end-use facility 'demand' (TJ) by gas day, state, and sector.

    Returns {tab_name: DataFrame} - one tab per state present in the
    data (sorted), followed by "Total" (all states summed) last. No
    per-facility rows are kept in any of them.
    """
    df = pd.read_csv(io.StringIO(flow_text))
    df.columns = [c.strip().lower() for c in df.columns]

    df["facility_type"] = df["facilityid"].map(facility_types)
    df = df[df["facility_type"].isin(FACILITY_TYPE_TO_SECTOR.keys())].copy()
    if df.empty:
        return {}

    df["sector"] = df["facility_type"].map(FACILITY_TYPE_TO_SECTOR)
    df["gas_date"] = pd.to_datetime(df["gasdate"]).dt.date
    df["demand_tj"] = pd.to_numeric(df["demand"], errors="coerce").fillna(0.0)
    df["state"] = df["state"].astype(str).str.strip().str.upper()

    tabs = {}
    for state in sorted(s for s in df["state"].unique() if s and s != "NAN"):
        tabs[state] = _pivot_daily_sector(df[df["state"] == state])
    tabs["Total"] = _pivot_daily_sector(df)
    return tabs


UNITS_SECTION_TITLES = {"UNITS", "SECTOR DEFINITIONS", "STATE COVERAGE", "IMPORTANT CAVEAT"}


def build_units_notes(state_tabs):
    lines = [
        "UNITS",
        "All four sector columns are in terajoules (TJ) per gas day, as published by AEMO - "
        "not converted to MW/MWh like the other scripts in this repo.",
        "",
        "SECTOR DEFINITIONS",
        "gas_powered_generation: AEMO facility type BBGPG",
        "large_industrial_users: AEMO facility type BBLARGE",
        "blended_distribution: AEMO facility type BDIST (residential/commercial city-gas)",
        "lng_export: AEMO facility type LNGEXPORT",
        "Pipeline, production, storage and compression facilities are excluded - summing "
        "those in would double-count gas that pipelines also report as \"demand\" simply "
        "because they're transporting it, not consuming it.",
        "",
        "STATE COVERAGE",
        f"State tabs in this run: {', '.join(state_tabs)}. State comes directly from AEMO's "
        "own \"State\" field on each flow row.",
        "Western Australia is NOT included - it runs a separate Gas Bulletin Board "
        "(gbbwa.aemo.com.au) covering a genuinely different, isolated system.",
        "AEMO's own registry has no separate ACT code - ACT facilities report under NSW, so "
        "there's no distinct \"ACT\" tab even though this covers the whole east coast + NT.",
        "The Total tab is the sum of every state tab (not a separately-published AEMO figure), "
        "so it equals whichever state tabs are actually present in this workbook.",
        "",
        "IMPORTANT CAVEAT",
        f"AEMO expanded the Bulletin Board's reporting scope on {GBB_EXPANSION_DATE.isoformat()}. "
        "Before that date, NONE of these four end-use facility types reported at all - it's a "
        "real reporting gap, not missing data. Any chart of this series across that boundary "
        "shows demand rising from zero, which is a reporting change, not a market event.",
        "The \"gas day\" (AEMO's GasDate) runs 06:00-06:00 fixed Australian Eastern Standard "
        "Time (UTC+10, no daylight saving) and is used directly as published.",
    ]
    return lines


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from-date", type=date.fromisoformat, default=date(2021, 10, 1))
    parser.add_argument("--to-date", type=date.fromisoformat, default=date.today())
    parser.add_argument("--out", default="australia_gas_demand_by_sector.xlsx")
    args = parser.parse_args()

    session = make_session()
    print("Fetching facility reference list...", file=sys.stderr)
    facility_types = fetch_facilities(session)
    print(f"  {len(facility_types)} facilities loaded", file=sys.stderr)

    print("Fetching full-history flow archive (this is a sizeable download)...", file=sys.stderr)
    flow_text = fetch_flow_archive(session)

    print("Aggregating to daily totals by state and sector (no per-facility rows kept)...", file=sys.stderr)
    tabs = aggregate_daily_sector_demand_by_state(flow_text, facility_types)

    filtered_tabs = {}
    for name, df in tabs.items():
        df = df[(df.index >= args.from_date) & (df.index <= args.to_date)]
        if not df.empty:
            filtered_tabs[name] = df

    if not filtered_tabs:
        print("No data returned for the requested range.", file=sys.stderr)
        sys.exit(1)

    state_names = [name for name in filtered_tabs if name != "Total"]
    xlsx_notes.write_workbook(args.out, filtered_tabs, build_units_notes(state_names), UNITS_SECTION_TITLES)
    print(f"Saved tabs {list(filtered_tabs.keys())} to {args.out}")
    if args.from_date < GBB_EXPANSION_DATE:
        print(
            f"Note: end-use sector reporting only begins {GBB_EXPANSION_DATE.isoformat()} - "
            "rows before that date will be all zero.",
            file=sys.stderr,
        )
    print(filtered_tabs["Total"].head())
    print(filtered_tabs["Total"].tail())


if __name__ == "__main__":
    main()
