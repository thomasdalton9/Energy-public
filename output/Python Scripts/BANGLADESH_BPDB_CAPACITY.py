"""
Bangladesh installed and derated generating capacity by fuel from BPDB (Bangladesh Power Development Board),
'Power Generation Units (Fuel Type Wise)': https://misc.bpdb.gov.bd/power-generation-unit - the page loads its
table from https://misc.bpdb.gov.bd/power-generation-unit-search?capacity_type=all (JSON, no key):
  installation  installed capacity by fuel (Gas, HFO, HSD, Coal, Imported, Solar, Hydro, Wind, none), MW
  present       derated (present) capacity, same fuels
Found via discovery_archive/asia/CAPPRICE_DISCOVERY1-3.py.

Writes output/Data and Chart Outputs/bangladesh_power_capacity.xlsx (standard capacity layout,
south_america/power_capacity_std.py):
  Monthly  installed MW by fuel, one row per month (the latest run in that month); Gas; Oil = HFO + HSD; Coal;
           Hydro; Solar; Wind; Other = BPDB's 'none'. Imports_MW (cross-border import capacity from India, BPDB
           'Imported') is kept outside Total_MW.
  Derated  the same for derated (present) capacity
  Latest   the latest raw table (fuel, installed MW, derated MW)

BPDB publishes only the current figures, so history starts with this pull's first run. Runs on the 1st and 15th.

    python3 asia/BANGLADESH_BPDB_CAPACITY.py [--out PATH]
"""
import argparse
import os
import sys
import time
from datetime import date

import pandas as pd
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import power_capacity_std as cap_std  # noqa: E402

urllib3.disable_warnings()   # BPDB's certificate chain is incomplete
PAGE = "https://misc.bpdb.gov.bd/power-generation-unit"
API = "https://misc.bpdb.gov.bd/power-generation-unit-search"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json, */*", "X-Requested-With": "XMLHttpRequest", "Referer": PAGE}
T = (20, 90)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "bangladesh_power_capacity.xlsx")
FUEL = {"gas": "Gas", "hfo": "Oil", "hsd": "Oil", "diesel": "Oil", "furnace": "Oil", "oil": "Oil", "coal": "Coal",
        "hydro": "Hydro", "solar": "Solar", "wind": "Wind", "imported": "Imports", "import": "Imports",
        "nuclear": "Nuclear", "biomass": "Bioenergy"}


def out(*a):
    print(*a, flush=True)


def fetch():
    for i in range(4):
        try:
            r = requests.get(API, params={"capacity_type": "all"}, headers=H, timeout=T, verify=False)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                raise
            out(f"  retry: {e}")
            time.sleep(5 * (i + 1))


def table(items):
    return {(x.get("fuel_type") or {}).get("title_en", str(x.get("fuel_type_id"))).strip():
            float(x.get("capacity") or 0) for x in items or []}


def row(by_fuel_raw, month):
    agg = {}
    for name, mw in by_fuel_raw.items():
        f = FUEL.get(name.lower(), "Other")
        agg[f] = agg.get(f, 0.0) + mw
    std = cap_std.standard(pd.DataFrame([{k: v for k, v in agg.items() if k != "Imports"}], index=[month]))
    std["Imports_MW"] = agg.get("Imports", 0.0)
    return std


def merge(old, new):
    if old.empty:
        return new
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    j = fetch()
    inst, pres = table(j.get("installation")), table(j.get("present"))
    out(f"Installed: {inst}\nDerated: {pres}")
    if sum(inst.values()) < 15000:
        raise SystemExit(f"Implausible installed total {sum(inst.values())} MW")
    month = pd.Timestamp(date.today().replace(day=1))
    monthly = merge(cap_std.load_monthly(a.out), row(inst, month))
    derated = cap_std.load_sheet(a.out, "Derated", index_col=0)
    if not derated.empty:
        derated.index = pd.to_datetime(derated.index)
    derated = merge(derated, row(pres, month)) if pres else derated
    monthly.index.name = derated.index.name = "date"
    latest = pd.DataFrame({"Installed_MW": pd.Series(inst), "Derated_MW": pd.Series(pres)})
    latest.index.name = "BPDB fuel type"
    latest["Standard_fuel"] = [FUEL.get(k.lower(), "Other") for k in latest.index]
    notes = [
        "UNITS",
        "Monthly: installed capacity, MW, by fuel, from BPDB's fuel-type-wise table on the run date (one row per month: "
        "the latest run in that month). Gas; Oil = HFO + HSD (diesel); Coal; Hydro; Solar; Wind; Other = BPDB 'none'. "
        "Total_MW = sum of fuels (domestic plant). Imports_MW = cross-border import capacity (BPDB 'Imported', from "
        "India), not in Total_MW.",
        "Derated: derated (present) capacity, same layout. Latest: BPDB's table as published on the latest run.",
        "",
        "COVERAGE",
        f"BPDB publishes only the current figures, so the history starts with this pull's first run "
        f"({monthly.index.min():%b %Y}). Grid-connected plants (public, IPP, rental and import) as BPDB counts them; "
        "captive generation and off-grid solar (solar home systems) are not included.",
        "",
        "SOURCE",
        f"BPDB (Bangladesh Power Development Board), Power Generation Units (Fuel Type Wise): {PAGE} "
        f"(JSON {API}?capacity_type=all).",
    ]
    cap_std.write(a.out, monthly, {"Derated": derated, "Latest": latest}, notes, {"UNITS", "COVERAGE", "SOURCE"})


if __name__ == "__main__":
    main()
