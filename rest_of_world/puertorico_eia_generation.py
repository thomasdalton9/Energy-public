"""
Pull Puerto Rico's monthly power generation from EIA's electric-power-
operational-data, broken down two ways: by sector (Electric Utility vs
Independent Power Producers vs Industrial/Commercial Combined Heat and
Power) and by fuel type - both more granular than a single "PR total"
number.

Data source: EIA API v2, electricity/electric-power-operational-data -
https://api.eia.gov/v2/electricity/electric-power-operational-data/data/
Needs a free EIA API key, read from the EIA_API_KEY environment
variable (a GitHub Actions secret in this repo) - same pattern as
EIA930_FUEL_MIX_DAILY.py.

SECTOR BREAKDOWN: confirmed via PUERTORICO_EIA_SECTOR_NAMES_INSPECT.py
that sectors 1-7 sum EXACTLY to sector 99 ("All Sectors") for PR, so
they're genuinely mutually exclusive - no double-counting risk:
    1 Electric Utility, 2 IPP Non-CHP, 3 IPP CHP, 4 Commercial Non-CHP,
    5 Commercial CHP, 6 Industrial Non-CHP, 7 Industrial CHP.

FUEL TYPE BREAKDOWN: EIA's fueltypeid facet mixes real leaf categories
with aggregates of each other (e.g. PET/PEL = DFO + RFO; COL/COW/BIS/
BIT are all identical for PR, since its coal is 100% bituminous; AOR/
REN = the renewables sum; NGO = NG, no separate "other gases" for PR).
Verified the exact arithmetic live (PUERTORICO_EIA_FUELTYPE_HIERARCHY_
INSPECT.py, sectorid=99, period 2026-07): the leaf set below sums
EXACTLY to the "ALL" total (1607.53651 = 304.718 COL + 191.93883 DFO +
413.97958 RFO + 648.63659 NG + 18.66351 SUN + 29.736 WND + (-0.136)
OTH, with HYC/BIO/WAS/LFG/MLG/OOG/GEO/NUC all 0 that period but kept
in case they ever report). AGGREGATE_FUELTYPE_IDS below is excluded
explicitly rather than hardcoding a fixed leaf list, so a genuinely new
leaf category EIA adds still shows up as its own column.
"""

import argparse
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

API_KEY = os.environ.get("EIA_API_KEY")
URL = "https://api.eia.gov/v2/electricity/electric-power-operational-data/data/"
TIMEOUT = (10, 45)
PAGE_LENGTH = 5000

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "puertorico_generation_monthly.xlsx")

SECTOR_NAMES = {
    "1": "Electric_Utility",
    "2": "IPP_Non_CHP",
    "3": "IPP_CHP",
    "4": "Commercial_Non_CHP",
    "5": "Commercial_CHP",
    "6": "Industrial_Non_CHP",
    "7": "Industrial_CHP",
}

# Aggregates of other fuel-type ids (verified live - see module docstring)
# excluded explicitly so any genuinely new leaf category EIA adds still
# shows up as its own column, rather than being silently summed twice.
AGGREGATE_FUELTYPE_IDS = {"ALL", "AOR", "REN", "FOS", "COW", "BIS", "BIT", "PEL", "PET", "NGO", "SPV", "WNT"}


def fetch_page(facets, offset):
    params = {
        "api_key": API_KEY,
        "frequency": "monthly",
        "data[]": "generation",
        "facets[location][]": "PR",
        "sort[0][column]": "period",
        "sort[0][direction]": "asc",
        "offset": offset,
        "length": PAGE_LENGTH,
        **facets,
    }
    r = requests.get(URL, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def fetch_all(facets):
    rows = []
    offset = 0
    while True:
        payload = fetch_page(facets, offset)
        page_rows = payload["response"]["data"]
        rows.extend(page_rows)
        total = int(payload["response"]["total"])
        offset += len(page_rows)
        if not page_rows or offset >= total:
            break
    return rows


def by_sector():
    rows = fetch_all({"facets[sectorid][]": list(SECTOR_NAMES.keys()), "facets[fueltypeid][]": "ALL"})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["generation"] = pd.to_numeric(df["generation"], errors="coerce")
    df["sector"] = df["sectorid"].map(SECTOR_NAMES)
    wide = df.pivot_table(index="period", columns="sector", values="generation", aggfunc="last")
    wide = wide[[c for c in SECTOR_NAMES.values() if c in wide.columns]]
    wide["Total_All_Sectors"] = wide.sum(axis=1, skipna=True)
    wide.index.name = "period"
    return wide.sort_index()


def by_fuel_type():
    rows = fetch_all({"facets[sectorid][]": "99"})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df[~df["fueltypeid"].isin(AGGREGATE_FUELTYPE_IDS)]
    df["generation"] = pd.to_numeric(df["generation"], errors="coerce")
    wide = df.pivot_table(index="period", columns="fueltypeid", values="generation", aggfunc="last")
    wide["Total_All_Fuels"] = wide.sum(axis=1, skipna=True)
    wide.index.name = "period"
    return wide.sort_index()


NOTES_LINES = [
    "UNITS",
    "All generation columns are in thousand megawatthours (GWh) for that month.",
    "",
    "TABS",
    "By sector: Electric_Utility, IPP_Non_CHP (independent power producers, non-cogeneration), "
    "IPP_CHP, Commercial_Non_CHP, Commercial_CHP, Industrial_Non_CHP, Industrial_CHP - confirmed to "
    "sum exactly to EIA's own 'All Sectors' total, so these are genuinely mutually exclusive.",
    "By fuel type: EIA's fueltypeid facet reported for sectorid=99 (All Sectors), with known "
    "aggregate ids (ALL, AOR/REN, FOS, COW/BIS/BIT, PEL/PET, NGO, SPV, WNT) excluded so nothing is "
    "double-counted - verified live that the remaining ids sum exactly to the 'ALL' total for a "
    "sample period. Common columns: COL (coal), DFO (distillate fuel oil), RFO (residual fuel oil), "
    "NG (natural gas), SUN (solar), WND (wind); HYC (hydro), BIO (biomass), and others appear only in "
    "periods/fuels Puerto Rico actually reports.",
    "Both tabs' Total column should match each other and EIA's own 'All Sectors'/'ALL' totals.",
    "",
    "SOURCE",
    f"EIA (US Energy Information Administration) API v2, electric-power-operational-data "
    f"(EIA-923-based): {URL}",
]
NOTES_SECTION_TITLES = {"UNITS", "TABS", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    if not API_KEY:
        print("EIA_API_KEY is not set in the environment - aborting.", file=sys.stderr)
        sys.exit(1)

    sector_df = by_sector()
    fuel_df = by_fuel_type()
    print(f"By sector: {sector_df.shape}", file=sys.stderr)
    print(sector_df.tail(3).to_string(), file=sys.stderr)
    print(f"\nBy fuel type: {fuel_df.shape}", file=sys.stderr)
    print(fuel_df.tail(3).to_string(), file=sys.stderr)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(
        args.out, {"By sector": sector_df, "By fuel type": fuel_df}, NOTES_LINES, NOTES_SECTION_TITLES
    )
    print(f"\nSaved to {args.out}")


if __name__ == "__main__":
    main()
