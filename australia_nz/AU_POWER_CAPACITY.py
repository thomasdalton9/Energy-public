"""
Australia installed generating capacity by fuel, from AEMO's registration
lists (public, no key):

  NEM  https://www.aemo.com.au/-/media/Files/Electricity/NEM/Participant_Information/NEM-Registration-and-Exemption-List.xls
       sheet "PU and Scheduled Loads": Region, Fuel Source / Technology Type descriptors, Reg Cap generation (MW)
  WEM  https://data.wa.aemo.com.au/public/public-data/datafiles/facilities/facilities.csv
       Facility Code, Facility Type, Maximum Capacity (MW) - fuel from the facility code (as in AU_WEM_GENERATION)

Both lists are snapshots of what is registered now, with no history, so each
run writes the current month's row and keeps the rows saved by earlier runs:
the series starts with the first run and builds up month by month.

Writes au_power_capacity.xlsx: "Monthly" (standard layout: date, <Fuel>_MW, Total_MW, plus
Battery_storage_MW and Pumped_storage_MW kept out of Total_MW) and "By region" (MW by NEM region and WEM,
same month rows).

Usage: python3 AU_POWER_CAPACITY.py [--out "output/Data and Chart Outputs/au_power_capacity.xlsx"]
"""
import argparse
import io
import os
import sys
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
import aemo_registration  # noqa: E402
from AU_WEM_GENERATION import fuel_of  # noqa: E402

WEM_FACILITIES = "https://data.wa.aemo.com.au/public/public-data/datafiles/facilities/facilities.csv"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_power_capacity.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Nuclear", "Other"]
STORAGE = ["Battery_storage", "Pumped_storage"]


def nem_units():
    return aemo_registration.units()


def wem_units():
    r = requests.get(WEM_FACILITIES, headers={"User-Agent": "Mozilla/5.0"}, timeout=(10, 120))
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content))
    tcol = next((c for c in d.columns if "type" in c.lower()), None)
    if tcol:
        d = d[d[tcol].astype(str).str.contains("Gen|Storage", case=False)]   # generators and storage, not loads
    capcol = next((c for c in d.columns if "maximum capacity" in c.lower()),
                  next(c for c in d.columns if "capacity" in c.lower() and "credit" not in c.lower()))
    fuel = d["Facility Code"].map(fuel_of).replace({"Battery_discharge": "Battery_storage"})
    out = pd.DataFrame({"region": "WA (WEM)", "fuel": fuel, "mw": pd.to_numeric(d[capcol], errors="coerce")})
    print(f"WEM: {len(out)} facilities, {out['mw'].sum() / 1000:.1f} GW ({capcol}); types "
          f"{d[tcol].value_counts().to_dict() if tcol else 'n/a'}", flush=True)
    return out


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    units = pd.concat([nem_units(), wem_units()], ignore_index=True)
    month = pd.Timestamp(date.today()).to_period("M").to_timestamp()
    by_fuel = units.groupby("fuel")["mw"].sum()
    row = pd.DataFrame([{f"{f}_MW": by_fuel.get(f) for f in FUELS if f in by_fuel}], index=[month])
    row["Total_MW"] = row.sum(axis=1)
    for s in STORAGE:
        if s in by_fuel:
            row[f"{s}_MW"] = by_fuel[s]
    reg = units[~units["fuel"].isin(STORAGE)].groupby("region")["mw"].sum().to_frame(month).T.add_suffix("_MW")

    monthly, region = load(args.out, "Monthly"), load(args.out, "By region")
    # this run's snapshot REPLACES any row saved earlier for the same month (no stale columns carried over)
    monthly = pd.concat([monthly.drop(index=month, errors="ignore"), row]).sort_index().dropna(axis=1, how="all")
    region = pd.concat([region.drop(index=month, errors="ignore"), reg]).sort_index().dropna(axis=1, how="all")
    monthly = monthly[[c for c in row.columns] + [c for c in monthly.columns if c not in row.columns]]
    for x in (monthly, region):
        x.index.name = "date"
    print(monthly.tail(3).to_string(), flush=True)
    notes = [
        "UNITS",
        "Registered generating capacity, MW, as registered on the date of each run (one row per month, the "
        "latest run in the month wins). NEM: 'Reg Cap generation (MW)' per unit; WEM: 'Maximum Capacity (MW)'.",
        "Monthly: Hydro (excl. pumped storage), Gas, Wind, Solar (utility scale), Coal, Oil, Bioenergy; "
        "Battery_storage_MW and Pumped_storage_MW are storage, kept out of Total_MW.",
        "NEM fuel = AEMO's Fuel Source / Technology Type descriptors; WEM fuel = from the facility code (AEMO's "
        "WEM list has no fuel field; see au_wem_power_generation_daily.xlsx Units for the rules).",
        "By region: generating capacity (excl. storage) by NEM region and WEM.",
        "",
        "COVERAGE",
        f"NEM (QLD, NSW, VIC, SA, TAS) + WA's WEM. Rooftop solar, NT and off-grid plant not included. Series "
        f"starts with the first run ({monthly.index.min():%b %Y}) - AEMO publishes no capacity history in these "
        "lists.",
        "",
        "SOURCE",
        f"AEMO NEM Registration and Exemption List: {aemo_registration.URL}; AEMO WA facilities list: "
        f"{WEM_FACILITIES}",
        "https://aemo.com.au/en/energy-systems/electricity/national-electricity-market-nem/participate-in-the-market/registration",
    ]
    xlsx_notes.write_workbook(args.out, {"Monthly": monthly.round(1), "By region": region.round(1)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
