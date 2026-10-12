"""
Ember yearly installed generating capacity by fuel for EVERY country, one whole-file download.
   output/Data and Chart Outputs/ember_capacity_yearly.xlsx

The fallback source for the complete monthly capacity builds (capacity_fill.py: res "EMBER", flag FALLBACK_EMBER) wherever a country
has no national capacity feed, and the reference series a national annual or late-starting source is spliced against. LONG_TERM_ENERGY.py
already carries Ember capacity but lumps oil, bioenergy and other renewables into one 'Other'; this keeps them apart, in the repo's
standard fuel columns, so later regions (Asia, Oceania, East Med, Africa) can use it too.

Source: Ember yearly electricity data (CC-BY-4.0), category 'Capacity', subcategory 'Fuel', GW:
  https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv
Mapping: Hydro, Gas, Coal, Nuclear, Wind, Solar, Bioenergy as Ember; Oil = Ember 'Other Fossil' (mostly oil-fired plant);
Other = 'Other Renewables' (geothermal, marine, ...). Ember has no battery or pumped-storage series. A fuel Ember does not list for a
country is left BLANK (not zero); a listed zero is kept as 0.

Sheet 'Yearly': iso3, country, year, <Fuel>_MW (year-END stock, MW), Total_MW. Dates: the year Y value is the stock at 31 December Y,
i.e. 'date_1_Jan' = 1 January Y+1 in the repo-wide capacity convention (a row dated the 1st = stock at the end of the previous month).
Usage: python3 EMBER_CAPACITY_YEARLY.py [--out FILE] [--csv LOCAL_CSV]
"""
import argparse
import io
import os

import pandas as pd
import requests

import xlsx_notes

ROOT = os.path.dirname(os.path.abspath(__file__))
URL = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv"
FUEL = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal", "Nuclear": "Nuclear",
        "Other Fossil": "Oil", "Bioenergy": "Bioenergy", "Other Renewables": "Other"}
COLS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
START_YEAR = 2015


def build(raw):
    d = raw[(raw["Category"] == "Capacity") & (raw["Subcategory"] == "Fuel") & (raw["Unit"] == "GW") & (raw["Year"] >= START_YEAR)]
    d = d[d["ISO 3 code"].astype(str).str.fullmatch(r"[A-Z]{3}")].copy()
    d["fuel"] = d["Variable"].map(FUEL)
    d = d.dropna(subset=["fuel"])
    d["MW"] = pd.to_numeric(d["Value"], errors="coerce") * 1000.0
    d = d.dropna(subset=["MW"])      # Ember lists some fuels with an empty value: that is 'not available', never 0
    w = d.pivot_table(index=["ISO 3 code", "Area", "Year"], columns="fuel", values="MW", aggfunc="sum")
    w = w.reindex(columns=COLS).add_suffix("_MW").reset_index()
    w.columns = ["iso3", "country", "year"] + list(w.columns[3:])
    w["Total_MW"] = w[[f"{c}_MW" for c in COLS]].sum(axis=1, min_count=1)
    w["date_1_Jan"] = pd.to_datetime((w["year"].astype(int) + 1).astype(str) + "-01-01")
    cols = ["iso3", "country", "year", "date_1_Jan"] + [f"{c}_MW" for c in COLS] + ["Total_MW"]
    w = w[cols].sort_values(["iso3", "year"]).reset_index(drop=True)
    num = [c for c in cols if c.endswith("_MW")]
    w[num] = w[num].round(1)
    return w


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "output", "Data and Chart Outputs", "ember_capacity_yearly.xlsx"))
    ap.add_argument("--csv", default=None, help="local copy of the Ember CSV (testing)")
    args = ap.parse_args()
    if args.csv:
        raw = pd.read_csv(args.csv)
    else:
        print(f"Downloading {URL}", flush=True)
        r = requests.get(URL, timeout=(10, 300))
        r.raise_for_status()
        raw = pd.read_csv(io.BytesIO(r.content))
    out = build(raw)
    print(f"{out['iso3'].nunique()} countries, {out['year'].min()}-{out['year'].max()}, {len(out)} rows", flush=True)
    notes = [
        "UNITS",
        "Installed generating capacity by fuel, MW (Ember publishes GW), one row per country and YEAR. The value is the year-END stock (31 December); "
        "'date_1_Jan' is 1 January of the NEXT year, the repo-wide capacity convention (a row dated the 1st = stock at the end of the previous month).",
        "Hydro, Gas, Coal, Nuclear, Wind, Solar, Bioenergy as Ember. Oil = Ember 'Other Fossil' (mostly oil-fired plant). Other = Ember 'Other Renewables' "
        "(geothermal and others). Ember has no battery or pumped-storage capacity. A fuel Ember does not list for a country is blank, a listed zero is 0.",
        "",
        "COVERAGE",
        f"{out['iso3'].nunique()} countries and areas, {out['year'].min()} to {out['year'].max()} (the latest year differs by country).",
        "",
        "SOURCE",
        f"Ember yearly electricity data (CC-BY-4.0), category Capacity: {URL}. Used as the labelled FALLBACK (flag FALLBACK_EMBER) in the "
        "complete monthly capacity workbooks where a country has no national capacity source, and as the splice reference for national annual sources.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Yearly": out.set_index("iso3")}, notes, ["UNITS", "COVERAGE", "SOURCE"])
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
