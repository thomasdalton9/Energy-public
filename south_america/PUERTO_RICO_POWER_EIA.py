"""
Puerto Rico power generation by type, monthly, in the standard raw-power
layout the South & Central America master reads.

Source: EIA (US Energy Information Administration) API v2,
electricity/electric-power-operational-data - the monthly EIA-923 survey
that every Puerto Rico plant (PREPA / Genera PR units and the independent
producers EcoEléctrica, AES Puerto Rico and the renewables) reports to.
  https://api.eia.gov/v2/electricity/electric-power-operational-data/data/
  location PR, sector 99 (all sectors), monthly, generation (thousand MWh)
Needs the EIA_API_KEY secret.

There is no public daily or hourly series for Puerto Rico by fuel (EIA-930
does not cover the island; LUMA / Genera PR publish only today's live
dashboard). The finest official granularity is MONTHLY, so sheet "Daily"
holds one row per month dated the first of the month with that month's
total MWh, and the notes say so.

Fuel mapping (EIA fueltypeid, leaf ids only - the aggregates ALL, FOS, PET,
PEL, COW, BIT, REN, AOR, NGO, SPV, WNT are left out so nothing is counted
twice; see rest_of_world/puertorico_eia_generation.py):
  NG -> Gas; COL -> Coal; RFO, DFO -> Oil; SUN -> Solar; WND -> Wind;
  HYC -> Hydro; BIO, LFG, WAS, WOO, MLG, OBW, OBS, OBL -> Bioenergy;
  everything else (OTH, OOG, PC, ...) -> Other.
Utility-scale only: EIA-923 does not include rooftop (distributed) solar.

Incremental: the first run fetches from 2021-01; later runs fetch only from
six months before the last saved month (EIA revises recent months) and merge.

Usage: python3 PUERTO_RICO_POWER_EIA.py [--out PATH] [--full]
"""

print("STARTING", flush=True)

import argparse
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import central_america_power_common as C  # noqa: E402

OUT = "output/Data and Chart Outputs/puerto_rico_power_generation_daily.xlsx"
URL = "https://api.eia.gov/v2/electricity/electric-power-operational-data/data/"
API_KEY = os.environ.get("EIA_API_KEY")
START = "2021-01"
REFRESH_MONTHS = 6

AGGREGATES = {"ALL", "AOR", "REN", "FOS", "COW", "BIS", "BIT", "PEL", "PET", "NGO", "SPV", "WNT", "SUB", "LIG"}
MAPPING = {"NG": "Gas", "COL": "Coal", "RFO": "Oil", "DFO": "Oil", "SUN": "Solar", "WND": "Wind", "HYC": "Hydro",
           "BIO": "Bioenergy", "LFG": "Bioenergy", "WAS": "Bioenergy", "WOO": "Bioenergy", "MLG": "Bioenergy",
           "OBW": "Bioenergy", "OBS": "Bioenergy", "OBL": "Bioenergy", "WWW": "Bioenergy"}


def fetch(start):
    rows, offset = [], 0
    while True:
        r = requests.get(URL, params={
            "api_key": API_KEY, "frequency": "monthly", "data[]": "generation", "facets[location][]": "PR",
            "facets[sectorid][]": "99", "start": start, "sort[0][column]": "period", "sort[0][direction]": "asc",
            "offset": offset, "length": 5000}, timeout=(10, 60))
        r.raise_for_status()
        resp = r.json()["response"]
        rows += resp["data"]
        offset += len(resp["data"])
        if not resp["data"] or offset >= int(resp["total"]):
            return pd.DataFrame(rows)


def to_standard(raw):
    raw = raw[~raw["fueltypeid"].isin(AGGREGATES)].copy()
    raw["MWh"] = pd.to_numeric(raw["generation"], errors="coerce") * 1000   # thousand MWh -> MWh
    raw["fuel"] = raw["fueltypeid"].map(MAPPING).fillna("Other")
    detail = raw.pivot_table(index="period", columns="fueltypeid", values="MWh", aggfunc="sum")
    fuels = raw.pivot_table(index="period", columns="fuel", values="MWh", aggfunc="sum")
    for f in (detail, fuels):
        f.index = pd.to_datetime(f.index)
        f.index.name = "date"
    detail.columns = [f"EIA_{c}_MWh" for c in detail.columns]
    std = C.standardise(fuels)
    keep = fuels.notna().any(axis=1)        # months EIA has published
    return std[keep.reindex(std.index).fillna(False)], detail


def notes(daily):
    last = daily.index.max()
    return [
        "UNITS",
        "MWh of net generation per MONTH. Puerto Rico has no public daily series by fuel, so each row is one "
        "month, dated the first of that month, holding the whole month's MWh (sheet name 'Daily' kept for the "
        "standard layout the master reads).",
        "Detail: the same data by EIA fuel code (MWh per month).",
        "",
        "COVERAGE",
        f"{daily.index.min():%b-%Y} to {last:%b-%Y}: {len(daily)} months. EIA publishes about two months after month "
        "end. Utility-scale plants only (EIA-923); rooftop solar is not included.",
        "",
        "SOURCE",
        "EIA (US Energy Information Administration), API v2 electricity/electric-power-operational-data "
        "(EIA-923 plant survey), location PR, all sectors: " + URL,
        "Browse: https://www.eia.gov/electricity/data/browser/ (State: Puerto Rico)",
        "",
        "MAPPING",
        "NG natural gas -> Gas; COL coal (AES Guayama) -> Coal; RFO residual fuel oil + DFO distillate -> Oil; "
        "SUN -> Solar; WND -> Wind; HYC -> Hydro; BIO/LFG/WAS/WOO/MLG -> Bioenergy; OTH and any other code -> Other. "
        "EIA aggregate codes (ALL, FOS, PET, PEL, COW, BIT, REN, AOR, NGO, SPV, WNT) are excluded so nothing is "
        "double counted.",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--full", action="store_true", help="re-fetch from 2021-01")
    args = ap.parse_args()
    if not API_KEY:
        sys.exit("EIA_API_KEY is not set")

    existing = C.load_sheet(args.out, "Daily")
    old_detail = C.load_sheet(args.out, "Detail")
    start = START
    if not args.full and not existing.empty:
        start = max(pd.Timestamp(START), existing.index.max() - pd.DateOffset(months=REFRESH_MONTHS)).strftime("%Y-%m")
    print(f"Fetching EIA-923 Puerto Rico generation from {start}", flush=True)
    raw = fetch(start)
    if raw.empty:
        print("EIA returned nothing new", flush=True)
        if existing.empty:
            sys.exit(1)
        return
    new, detail = to_standard(raw)
    daily = C.merge(new, existing)
    detail = C.merge(detail, old_detail)
    daily = daily[daily.index >= pd.Timestamp(START)]
    C.print_mapping({k: v for k, v in MAPPING.items() if k in set(raw["fueltypeid"])}, "EIA fueltypeid")
    C.write(args.out, daily, notes(daily), detail)
    print((daily.tail(6) / 1000).round(1).to_string(), flush=True)


if __name__ == "__main__":
    main()
