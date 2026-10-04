"""
Sri Lanka installed generating capacity by fuel from PUCSL's GenData platform (Public Utilities Commission of
Sri Lanka; plant data from the CEB system control centre): https://gendata.pucsl.gov.lk/api/metadata/power-plants
(no key) - the same plant list asia/SRI_LANKA_PUCSL.py uses to map dispatch to fuels; each plant carries its
capacity (MW).

Writes output/Data and Chart Outputs/sri_lanka_power_capacity.xlsx in the standard capacity layout
(south_america/power_capacity_std.py):
  Monthly  installed MW by fuel, one row per month (the latest run in the month). The list holds only today's
           fleet, so history starts with the first run and grows by one row a month.
  Plants   the latest list with the fuel each plant is counted under

Fuel map: Major / Mini hydro -> Hydro; Coal; Oil (CEB) and Oil (IPP) -> Oil (incl. the Kerawalapitiya and
Kelanitissa combined-cycle plants, which burn oil until LNG arrives); Wind; Solar (ground-mounted, estimated and
rooftop); Biomass and waste heat -> Bioenergy; BESS -> left out (storage). Runs on the 1st and 15th.

    python3 asia/SRI_LANKA_CAPACITY.py [--out PATH]
"""
import argparse
import os
import sys
import time
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import power_capacity_std as cap_std  # noqa: E402

URL = "https://gendata.pucsl.gov.lk/api/metadata/power-plants"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"}
T = (15, 180)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "sri_lanka_power_capacity.xlsx")
# keyword in energy type / technology / fuel (lower case) -> standard fuel; first match wins (None = left out)
FUEL_KEYS = [("bess", None), ("battery", None), ("hydro", "Hydro"), ("coal", "Coal"), ("wind", "Wind"),
             ("solar", "Solar"), ("biomass", "Bioenergy"), ("dendro", "Bioenergy"), ("bio", "Bioenergy"),
             ("waste", "Bioenergy"), ("lng", "Gas"), ("natural gas", "Gas"), ("oil", "Oil"), ("diesel", "Oil"),
             ("naphtha", "Oil"), ("fuel", "Oil"), ("thermal", "Oil"), ("combined cycle", "Oil")]

DUP_TYPES = r"telemetered|estimated|est\.|^ncre"   # aggregate / estimate entries duplicating listed plants


def out(*a):
    print(*a, flush=True)


def plants():
    for i in range(4):
        try:
            r = requests.get(URL, headers=H, timeout=T)
            r.raise_for_status()
            data = r.json().get("data", [])
            break
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                raise
            out(f"  retry: {e}")
            time.sleep(5 * (i + 1))
    rows = []
    for p in data:
        cx = p.get("powerPlantComplex") or {}
        et = cx.get("energyType") or {}
        text = " ".join(str(x) for x in (et.get("name"), et.get("slug"), p.get("technology"), p.get("fuelUsed"),
                                         cx.get("name"), p.get("name"))).lower()
        fuel = next((f for k, f in FUEL_KEYS if k in text), "Other")
        rows.append({"id": p.get("id"), "name": p.get("name"), "complex": cx.get("name"), "energy_type": et.get("name"),
                     "technology": p.get("technology"), "fuel_used": p.get("fuelUsed"),
                     "capacity_MW": pd.to_numeric(p.get("capacityData"), errors="coerce"),
                     "show_capacity": p.get("showCapacity"),
                     "fuel": fuel if fuel else "Storage (not counted)"})
    return pd.DataFrame(rows).drop_duplicates("id", keep="first")   # the list repeats some plant ids


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    pl = plants()
    if pl.empty or pl["capacity_MW"].sum() < 1000:
        raise SystemExit(f"Plant list empty or implausible ({pl['capacity_MW'].sum() if len(pl) else 0} MW)")
    out(pl.groupby(["energy_type", "fuel"])["capacity_MW"].agg(["sum", "count"]).to_string())
    # the list mixes physical plants with the aggregate entries the dispatch feed uses for the same fleet
    # (Mini Hydro 422 MW vs 'Minihydro Telemetered' 421 MW, NCRE Wind 163 MW vs Wind CEB 163 MW, rooftop solar
    # by tariff class vs 'Solar Roof Top PV (export) Est.'): the aggregates are listed but not counted
    pl["counted"] = ~pl["energy_type"].fillna("").str.contains(DUP_TYPES, case=False, regex=True) \
        & pl["fuel"].isin(cap_std.FUELS)
    out("showCapacity vs counted:\n" + pl.groupby(["show_capacity", "counted"])["capacity_MW"].agg(["sum", "count"]).to_string())
    out("Not counted:\n" + pl.loc[~pl["counted"] & (pl["capacity_MW"] > 0),
                                  ["name", "energy_type", "capacity_MW"]].to_string())
    by = pl[pl["counted"]].groupby("fuel")["capacity_MW"].sum()
    month = pd.Timestamp(date.today().replace(day=1))
    row = cap_std.standard(pd.DataFrame([by.to_dict()], index=[month]))
    monthly = cap_std.load_monthly(a.out)
    monthly = pd.concat([monthly[~monthly.index.isin(row.index)], row]).sort_index() if not monthly.empty else row
    monthly.index.name = "date"
    notes = [
        "UNITS",
        "Monthly: installed capacity, MW, by fuel, from the PUCSL GenData plant list on the run date (one row per "
        "month: the latest run in that month). Total_MW = sum of fuels (battery storage not counted).",
        "The plant list also holds aggregate entries the dispatch feed uses for the same fleet (energy types "
        "'Mini Hydro (Telemetered)', '... (Estimated)', 'Solar Roof Top PV (export) Est.', 'NCRE Wind'); they "
        "duplicate listed plants and are not counted (Plants sheet: counted = False).",
        "Plants: the latest list, with the standard fuel each plant is counted under.",
        "",
        "COVERAGE",
        f"PUCSL publishes only the current fleet, so the monthly history starts with this pull's first run "
        f"({monthly.index.min():%b %Y}). Includes CEB and IPP plants, mini hydro, wind, ground-mounted solar and "
        "PUCSL's rooftop-solar capacity entries. Oil includes the Kerawalapitiya and Kelanitissa combined-cycle "
        "plants (oil-fired; listed under Oil (CEB/IPP)).",
        "",
        "SOURCE",
        "PUCSL (Public Utilities Commission of Sri Lanka) GenData, power plant metadata: "
        "https://gendata.pucsl.gov.lk/ (API https://gendata.pucsl.gov.lk/api/metadata/power-plants).",
    ]
    cap_std.write(a.out, monthly, {"Plants": pl}, notes, {"UNITS", "COVERAGE", "SOURCE"})


if __name__ == "__main__":
    main()
