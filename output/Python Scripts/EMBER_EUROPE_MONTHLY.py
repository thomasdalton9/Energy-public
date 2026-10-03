"""
Monthly generation by fuel, demand and net imports for the European countries whose ENTSO-E feed is incomplete or
too short for the Europe master (Ireland, Albania), from Ember's free monthly data release (no key, CC-BY-4.0):
  https://ember-energy.org/data/monthly-electricity-data/
Ember compiles each country's grid-operator / ministry data. It is the labelled FALLBACK source (CLAUDE.md): Ireland's
ENTSO-E feed covers only about 46-77% of EirGrid's all-island demand and Albania's starts in May 2026.

  output/Data and Chart Outputs/ember_europe_power_monthly.xlsx
    one sheet per country: Month (YYYY-MM), <Fuel>_GWh (Ember fuel groups), Total_GWh, Demand_GWh, NetImports_GWh
    (positive = import), Units

Ember lags real time by a few months; each run re-downloads the single release file (about a minute).

Usage: python3 EMBER_EUROPE_MONTHLY.py [--countries Ireland,Albania] [--out FILE] [--start-date 2021-01-01]
"""
import argparse
import io
import os
import sys

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

URL = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/monthly_full_release_long_format.csv"
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs", "ember_europe_power_monthly.xlsx")


def table(d, country, start):
    gen = d[(d["Area"] == country) & (d["Category"] == "Electricity generation") & (d["Subcategory"] == "Fuel")
            & (d["Unit"] == "TWh")]
    if gen.empty:
        return None
    w = gen.pivot_table(index="Date", columns="Variable", values="Value", aggfunc="sum")
    w.index = pd.to_datetime(w.index)
    out = w.mul(1000).round(1)
    out.columns = [f"{c}_GWh" for c in out.columns]
    out["Total_GWh"] = out.sum(axis=1)
    for cat, var, col in (("Electricity demand", "Demand", "Demand_GWh"), ("Electricity imports", "Net imports", "NetImports_GWh")):
        x = d[(d["Area"] == country) & (d["Category"] == cat) & (d["Variable"] == var) & (d["Unit"] == "TWh")]
        if x.empty:
            print(f"  {country}: no {cat} / {var} rows", flush=True)
            continue
        s = x.groupby("Date")["Value"].sum()
        s.index = pd.to_datetime(s.index)
        out[col] = (s * 1000).round(1)
    out = out[out.index >= pd.Timestamp(start)].sort_index()
    out.index = out.index.strftime("%Y-%m")
    out.index.name = "Month"
    return out.reset_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--countries", default="Ireland,Albania")
    ap.add_argument("--out", default=OUT_DEFAULT)
    ap.add_argument("--start-date", default="2021-01-01")
    args = ap.parse_args()
    print(f"Downloading {URL}", flush=True)
    r = requests.get(URL, timeout=(10, 300))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    print("categories:", sorted(d["Category"].dropna().unique()), flush=True)
    sheets, cov = {}, []
    for c in [x.strip() for x in args.countries.split(",") if x.strip()]:
        t = table(d, c, args.start_date)
        if t is None:
            print(f"No Ember generation rows for {c!r} - skipped", flush=True)
            continue
        sheets[c] = t
        cov.append(f"{c}: {t['Month'].min()} to {t['Month'].max()}")
        print(f"{c}: {len(t)} months to {t['Month'].max()}; columns {list(t.columns)}", flush=True)
        print(t.tail(2).to_string(index=False), flush=True)
    if not sheets:
        raise SystemExit("No Ember data for any requested country")
    notes = ["UNITS", "GWh per month (Ember publishes TWh). Demand_GWh = electricity demand; NetImports_GWh = net imports "
             "(positive = import). Fuel groups are Ember's (Bioenergy, Coal, Gas, Hydro, Nuclear, Other Fossil, "
             "Other Renewables, Solar, Wind).", "",
             "COVERAGE", "; ".join(cov) + ". Ember lags real time by one to several months.", "",
             "SOURCE", f"Ember monthly electricity data (free, CC-BY-4.0): {URL}. Ember compiles grid operator / ministry data "
             "(Ireland: EirGrid / SONI all-island; Albania: OST). Fallback for countries whose ENTSO-E feed is incomplete."]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
