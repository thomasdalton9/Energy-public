"""
Monthly electricity generation by fuel/technology from Ember's free
monthly data release (no key). Ember compiles each country's grid
operator data - for Chile that's the Coordinador Electrico Nacional
(CEN), whose own API needs a registered key. Lags real time by a few
months.
  https://ember-energy.org/data/monthly-electricity-data/

Also writes an ESTIMATE of gas burned for power: gas-fired generation
divided by an assumed average fleet efficiency (50%, typical for a
combined-cycle-heavy fleet), converted to million m3/day at 38 MJ/m3.
It's a proxy for the power sector's gas demand where no published
monthly gas-by-sector split exists (e.g. Chile).

Usage: python3 EMBER_POWER_BY_TYPE.py --country Chile --out "output/Data and Chart Outputs/chile_power_by_type.xlsx"
"""
import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

URL = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/monthly_full_release_long_format.csv"
DATA_START = "2021-01-01"
EFFICIENCY = 0.50          # assumed average gas fleet efficiency (electric out / fuel in)
MJ_PER_M3 = 38.0           # natural gas gross heating value


def country_table(d, country, start_date):
    c = d[(d["Area"] == country) & (d["Category"] == "Electricity generation")
          & (d["Subcategory"] == "Fuel") & (d["Unit"] == "TWh")]
    if c.empty:
        return None
    w = c.pivot_table(index="Date", columns="Variable", values="Value", aggfunc="sum")
    w.index = pd.to_datetime(w.index)
    w = w[w.index >= start_date].sort_index()
    w["Total"] = w.sum(axis=1)
    out = w.mul(1000).round(1)  # GWh
    out.columns = [f"{col}_GWh" for col in out.columns]
    if "Gas" in w.columns:
        days = w.index.days_in_month
        fuel_mj = w["Gas"] * 1e9 * 3.6 / EFFICIENCY          # TWh_e -> MJ of fuel
        out["Gas_for_power_est_mcm_per_day"] = (fuel_mj / MJ_PER_M3 / 1e6 / days).round(2)
    out.index.name = "Month"
    out = out.reset_index()
    out["Month"] = out["Month"].dt.strftime("%Y-%m")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--country", default="Chile", help="one country -> sheet 'Generation by type'")
    ap.add_argument("--countries", default=None,
                    help="comma-separated list -> one sheet per country (e.g. the South America set)")
    ap.add_argument("--out", default="output/Data and Chart Outputs/chile_power_by_type.xlsx")
    ap.add_argument("--start-date", default=DATA_START)
    args = ap.parse_args()

    print(f"Downloading {URL}", flush=True)
    r = requests.get(URL, timeout=(10, 300))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    sheets, coverage = {}, []
    countries = [c.strip() for c in args.countries.split(",")] if args.countries else [args.country]
    for country in countries:
        out = country_table(d, country, args.start_date)
        if out is None:
            print(f"No Ember generation rows for {country!r} - skipped", flush=True)
            continue
        name = "Generation by type" if not args.countries else country[:31]
        sheets[name] = out
        coverage.append(f"{country}: {out['Month'].min()} to {out['Month'].max()}")
        print(f"{country}: {len(out)} months to {out['Month'].max()}", flush=True)
    if not sheets:
        raise SystemExit("No Ember data for any requested country")
    print(next(iter(sheets.values())).tail(3).to_string(index=False), flush=True)

    notes = [
        "UNITS",
        "Generation in GWh per month (Ember publishes TWh; multiplied by 1,000).",
        "Gas_for_power_est_mcm_per_day: ESTIMATED gas burned by gas-fired plants, million m3 per day = gas "
        f"generation / {EFFICIENCY:.0%} assumed fleet efficiency, at {MJ_PER_M3:.0f} MJ/m3. A proxy, not a "
        "published figure: actual efficiency varies with the plant mix (open-cycle and diesel-capable units "
        "are less efficient), so treat the level as approximate and the month-to-month shape as the signal.",
        "",
        "COVERAGE",
        "Monthly. " + "; ".join(coverage) + ". Ember lags real time by one to several months (it varies by "
        "country); the weekly run picks up new months as they're released.",
        "",
        "SOURCE",
        f"Ember monthly electricity data (free, CC-BY-4.0): {URL}. Ember compiles each country's grid operator / "
        "ministry data (e.g. Chile: Coordinador Electrico Nacional; Brazil: ONS; Colombia: XM).",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
