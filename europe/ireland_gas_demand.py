"""
Pull Ireland's daily gas demand by sector from Gas Networks Ireland's
open data CSV, split into:

  Daily Metered (DM): large industrial sites metered daily, excluding
    power generation.
  Non Daily Metered (NDM): residential and small commercial sites
    (metered periodically, GNI models their daily demand).
  Large Daily Metered (LDM) non Power Gen: very large industrial sites,
    metered daily, excluding power generation.
  Power Generation: gas-fired power stations' own demand.
  Total ROI demand: Republic of Ireland total (sum of the above) -
    Northern Ireland is a separate network, not included.

Data source: data.gov.ie's CKAN API confirmed the real resource URL
(IRELAND_TURKEY_DISCOVERY.py) - a single evergreen CSV (not one file
per quarter), daily since 2018-01-01:
    https://data.gov.ie/api/3/action/package_show?id=dailygasdemandireland
    -> resources[] -> the CSV-format entry's "url"

The CSV's own filename embeds a year-month that changes as Gas Networks
Ireland republishes it (confirmed 2026-08 as of this script's writing)
so the URL is re-resolved via the CKAN API on every run rather than
hardcoded, to survive that renaming.
"""

import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

CKAN_PACKAGE_URL = "https://data.gov.ie/api/3/action/package_show"
DATASET_ID = "dailygasdemandireland"
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 45)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "ireland_gas_demand_daily.xlsx")

COLUMN_RENAME = {
    "Date": "date",
    "Daily Metered ": "DM_GWh",
    "Non Daily Metered (NDM)": "NDM_GWh",
    "Large Daily Metered (LDM) non Power Gen": "LDM_ex_PowerGen_GWh",
    "Power Generation": "PowerGen_GWh",
    "Total ROI demand": "Total_ROI_GWh",
}


def resolve_csv_url():
    r = requests.get(CKAN_PACKAGE_URL, headers=HEADERS, params={"id": DATASET_ID}, timeout=TIMEOUT)
    r.raise_for_status()
    resources = r.json()["result"]["resources"]
    for res in resources:
        if (res.get("format") or "").upper() == "CSV":
            return res["url"]
    raise RuntimeError(f"No CSV resource found in data.gov.ie package {DATASET_ID!r}")


def fetch():
    csv_url = resolve_csv_url()
    print(f"Resolved CSV URL: {csv_url}", file=sys.stderr)
    r = requests.get(csv_url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    df = df.rename(columns=COLUMN_RENAME)
    missing = set(COLUMN_RENAME.values()) - set(df.columns)
    if missing:
        raise RuntimeError(f"Expected columns missing from the CSV - GNI may have changed its format: {missing}")
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df = df.set_index("date").sort_index()
    return df[["NDM_GWh", "LDM_ex_PowerGen_GWh", "DM_GWh", "PowerGen_GWh", "Total_ROI_GWh"]]


NOTES_LINES = [
    "UNITS",
    "All columns are in GWh (gigawatt-hours) of gas demand for that day.",
    "",
    "CATEGORIES",
    "NDM_GWh: Non Daily Metered - residential and small commercial, modelled (not directly metered "
    "daily) demand.",
    "LDM_ex_PowerGen_GWh: Large Daily Metered, excluding power generation - very large industrial "
    "sites.",
    "DM_GWh: Daily Metered - other large industrial sites metered daily, excluding power generation.",
    "PowerGen_GWh: gas-fired power stations' own gas demand.",
    "Total_ROI_GWh: Republic of Ireland total (sum of the above). Northern Ireland is a separate gas "
    "network, not included.",
    "",
    "COVERAGE",
    "Daily since 2018-01-01.",
    "",
    "SOURCE",
    "Gas Networks Ireland's open data CSV, resolved via data.gov.ie's CKAN API "
    f"(package id {DATASET_ID!r}) rather than a hardcoded URL, since GNI's own filename embeds a "
    "year-month that changes each time they republish the file.",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    df = fetch()
    print(f"{len(df)} days, {df.index.min()} to {df.index.max()}", file=sys.stderr)
    print(df.tail(10).to_string(), file=sys.stderr)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": df}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved to {args.out}")


if __name__ == "__main__":
    main()
