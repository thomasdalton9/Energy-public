"""
Pull Ireland's daily gas supply by source from Gas Networks Ireland's
open data CSV: indigenous production (Corrib, plus Inch/Kinsale until
it ceased in 2020) versus imports from Great Britain through the Moffat
interconnector - the supply-side counterpart of ireland_gas_demand.py.

Data source: data.gov.ie's CKAN package "dailygassupply" (daily since
2018-01-01, republished quarterly by GNI). As with the demand file, the
CSV's own filename embeds the quarter it was published (e.g.
".../2024-Q3-Daily-Gas-Supply.csv"), so the URL is re-resolved via the
CKAN API on every run rather than hardcoded.

Column names are matched by keyword rather than exact header text,
since the exact headers were not inspectable from the editing sandbox
(data.gov.ie is blocked by its egress proxy) - the run log prints the
raw headers so any drift is visible.
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
DATASET_ID = "dailygassupply"
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 45)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "ireland_gas_supply_daily.xlsx")

# (keywords that must ALL appear in the lower-cased header, output column)
COLUMN_RULES = [
    (("corrib",), "Corrib_Production_GWh"),
    (("inch",), "Inch_Production_GWh"),
    (("moffat",), "Moffat_Imports_GWh"),
    (("total",), "Total_Supply_GWh"),
]


def resolve_csv_url():
    r = requests.get(CKAN_PACKAGE_URL, headers=HEADERS, params={"id": DATASET_ID}, timeout=TIMEOUT)
    r.raise_for_status()
    for res in r.json()["result"]["resources"]:
        if (res.get("format") or "").upper() == "CSV":
            return res["url"]
    raise RuntimeError(f"No CSV resource found in data.gov.ie package {DATASET_ID!r}")


def map_columns(columns):
    mapping = {}
    for col in columns:
        low = col.lower()
        if low.strip() == "date" or low.startswith("date"):
            mapping[col] = "date"
            continue
        for keywords, out in COLUMN_RULES:
            if all(k in low for k in keywords) and out not in mapping.values():
                mapping[col] = out
                break
    return mapping


def fetch():
    csv_url = resolve_csv_url()
    print(f"Resolved CSV URL: {csv_url}", file=sys.stderr)
    r = requests.get(csv_url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    print(f"Raw headers: {list(df.columns)}", file=sys.stderr)
    mapping = map_columns(df.columns)
    df = df.rename(columns=mapping)
    if "date" not in df.columns or "Moffat_Imports_GWh" not in df.columns:
        raise RuntimeError(f"Could not identify the date/Moffat columns - headers were {list(df.columns)}")
    keep = ["date"] + [out for _, out in COLUMN_RULES if out in df.columns]
    df = df[keep].copy()
    # GNI writes month-first dates (confirmed live: parsing day-first
    # silently dropped every day 13-31 as unparseable and swapped the
    # rest). Parse both ways and keep whichever loses fewer rows, then
    # refuse to continue if more than a handful still failed.
    raw = df["date"].astype(str).str.strip()
    month_first = pd.to_datetime(raw, dayfirst=False, errors="coerce")
    day_first = pd.to_datetime(raw, dayfirst=True, errors="coerce")
    parsed = month_first if month_first.notna().sum() >= day_first.notna().sum() else day_first
    print(f"Date parse: month-first ok={month_first.notna().sum()} day-first ok={day_first.notna().sum()} "
          f"of {len(raw)} rows", file=sys.stderr)
    if parsed.isna().sum() > 5:
        raise RuntimeError(f"{parsed.isna().sum()} of {len(raw)} dates failed to parse - sample: "
                           f"{raw[parsed.isna()].head().tolist()}")
    df["date"] = parsed
    df = df.dropna(subset=["date"]).set_index("date").sort_index()
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df.index = df.index.date
    df.index.name = "date"
    return df


NOTES_LINES = [
    "UNITS",
    "All columns are in GWh of gas supplied to the Republic of Ireland network that day, as published by "
    "Gas Networks Ireland.",
    "",
    "SOURCES OF SUPPLY",
    "Corrib_Production_GWh: indigenous production from the Corrib field (via the Bellanaboy terminal).",
    "Inch_Production_GWh: indigenous production from the Kinsale/Seven Heads fields via the Inch entry point "
    "(ceased in 2020) - present only if the published file carries it.",
    "Moffat_Imports_GWh: imports from Great Britain through the Moffat interconnector (Scotland to Ireland, "
    "one-directional). Northern Ireland and the Isle of Man are also served via Moffat but are not part of "
    "this ROI figure.",
    "Total_Supply_GWh: GNI's published total - should track Total_ROI_GWh in ireland_gas_demand_daily.xlsx "
    "(the same network, supply side).",
    "",
    "COVERAGE",
    "Daily since 2018-01-01. GNI republishes the file quarterly, so the latest few months lag.",
    "",
    "SOURCE",
    f"Gas Networks Ireland open data CSV, resolved via data.gov.ie's CKAN API (package id {DATASET_ID!r}) "
    "rather than a hardcoded URL, since GNI's filename embeds the publication quarter.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCES OF SUPPLY", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    df = fetch()
    print(f"{len(df)} days, {df.index.min()} to {df.index.max()}", file=sys.stderr)
    print(df.tail(10).to_string(), file=sys.stderr)
    if "Total_Supply_GWh" in df.columns:
        share = df["Moffat_Imports_GWh"].iloc[-365:].sum() / df["Total_Supply_GWh"].iloc[-365:].sum() * 100
        print(f"Import share of supply, last 365 days: {share:.1f}%", file=sys.stderr)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": df}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved to {args.out}")


if __name__ == "__main__":
    main()
