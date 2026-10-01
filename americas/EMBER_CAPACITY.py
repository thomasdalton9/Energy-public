"""
Installed generating capacity by fuel, annual, from Ember's yearly
electricity data (free, CC-BY-4.0) - a FALLBACK for countries with no raw
capacity feed in this repo (Mexico: SENER/CENACE publish capacity only in
PDF reports such as PRODESEN).

  https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv

Writes the standard capacity layout (sheet "Monthly": date, <Fuel>_MW,
Total_MW) with one row per YEAR dated 1 January. Ember's "Other Fossil"
(mostly oil-fired plants) maps to Oil, "Other Renewables" (geothermal,
...) to Other.

Usage: python3 EMBER_CAPACITY.py --country Mexico --out "output/Data and Chart Outputs/mexico_power_capacity.xlsx"
"""
import argparse
import io
import os
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv"
START_YEAR = 2015
FUEL = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal", "Nuclear": "Nuclear",
        "Other Fossil": "Oil", "Bioenergy": "Bioenergy", "Other Renewables": "Other"}
COLS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--country", default="Mexico")
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "Data and Chart Outputs", "mexico_power_capacity.xlsx"))
    args = ap.parse_args()

    print(f"Downloading {URL}", flush=True)
    r = requests.get(URL, timeout=(10, 300))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    c = d[(d["Area"] == args.country) & (d["Category"] == "Capacity") & (d["Subcategory"] == "Fuel")
          & (d["Unit"] == "GW") & (d["Year"] >= START_YEAR)]
    if c.empty:
        raise SystemExit(f"No Ember capacity rows for {args.country!r}")
    print(f"Ember variables: {sorted(c['Variable'].unique())}", flush=True)
    c = c.assign(fuel=c["Variable"].map(FUEL).fillna("Other"))
    w = c.pivot_table(index="Year", columns="fuel", values="Value", aggfunc="sum") * 1000.0   # GW -> MW
    w = w.reindex(columns=[x for x in COLS if x in w.columns])
    out = w.add_suffix("_MW")
    out["Total_MW"] = w.sum(axis=1)
    out.index = pd.to_datetime(out.index.astype(int).astype(str) + "-01-01")
    out.index.name = "date"
    out = out.round(1)
    print(out.tail(3).to_string(), flush=True)

    notes = [
        "UNITS",
        "Installed generating capacity, MW (Ember publishes GW), one row per YEAR dated 1 January (year-end value).",
        "Oil = Ember 'Other Fossil' (mostly oil-fired plants); Other = 'Other Renewables' (geothermal and others).",
        "",
        "COVERAGE",
        f"{args.country}, annual, {out.index.min():%Y} to {out.index.max():%Y}.",
        "",
        "SOURCE",
        f"Ember yearly electricity data (free, CC-BY-4.0): {URL}. FALLBACK: used because this repo has no raw "
        f"capacity feed for {args.country} (Ember compiles it from the national energy ministry and grid operator).",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Monthly": out}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
