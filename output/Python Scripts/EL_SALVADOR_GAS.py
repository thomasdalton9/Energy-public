"""
El Salvador monthly natural-gas (LNG) use for power - ESTIMATE.

El Salvador's only gas consumer is the 378 MW Energia del Pacifico (EDP)
LNG-to-power plant at Acajutla (Wartsila 50DF engines, FSRU, commercial
operation 2022). No source reachable from GitHub Actions publishes its gas
burn or the country's LNG imports by month:
  - UT (Unidad de Transacciones, ut.com.sv) - the grid operator's servers do
    not answer outside El Salvador (name servers and web server time out);
  - SIGET, DGEHM (successor of CNE; cne.gob.sv no longer exists) and BCR
    publish no monthly LNG import / gas-burn series
    (see discovery_archive/south_america/NORTH_CENTRAL_AMERICA_POWER_PROBE3..12.py).
So gas use is estimated from monthly gas-fired generation x a stated heat
rate:
  gas generation : Ember monthly electricity data, El Salvador, 'Gas' (TWh),
                   which Ember compiles from UT - used because UT itself
                   cannot be reached; Ember rounds to 0.01 TWh (10 GWh);
  heat rate      : 8.2 MMBtu (HHV) per MWh, i.e. ~46% net LHV efficiency,
                   typical for Wartsila 50DF engines on gas at part load.
Conversions: 1 mcm of regasified LNG ~ 36,300 MMBtu (HHV 38.3 MJ/m3);
1 tonne LNG ~ 51.7 MMBtu (HHV).

Writes output/Data and Chart Outputs/el_salvador_gas.xlsx, sheet "Gas use".
Usage: python3 EL_SALVADOR_GAS.py [--out PATH] [--heat-rate MMBTU_PER_MWH]
"""

print("STARTING", flush=True)

import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/el_salvador_gas.xlsx"
EMBER = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/monthly_full_release_long_format.csv"
HEAT_RATE = 8.2            # MMBtu (HHV) per MWh
MMBTU_PER_MCM = 36_300.0   # HHV, regasified LNG
MMBTU_PER_T_LNG = 51.7     # HHV
START = "2021-01-01"


def notes(heat_rate):
    return [
        "UNITS",
        "Gas_generation_GWh: gas-fired electricity generation, GWh per month.",
        "Gas_use_MMBtu: ESTIMATED gas burned, MMBtu (HHV) per month = generation x heat rate. "
        "Gas_use_MMBtu_per_day: the same per calendar day. Gas_use_mcm: million cubic metres (regasified). "
        "LNG_equivalent_kt: thousand tonnes of LNG.",
        "",
        "METHOD (ESTIMATE)",
        f"Heat rate {heat_rate} MMBtu (HHV) per MWh (~46% net LHV efficiency), typical of the Wartsila 50DF engines at "
        "Energia del Pacifico (378 MW, Acajutla, El Salvador's only gas-fired plant, LNG via FSRU, commercial operation "
        "2022). Conversions: 36,300 MMBtu per mcm, 51.7 MMBtu per tonne of LNG (HHV).",
        "This is not a measured gas-use series: no source reachable from the pull (UT, SIGET, DGEHM, BCR) publishes El "
        "Salvador's monthly LNG imports or gas burn. Replace it if one appears.",
        "",
        "SOURCE",
        "Gas generation: Ember monthly electricity data (CC-BY-4.0), El Salvador, Category 'Electricity generation', "
        f"Variable 'Gas', {EMBER}. Ember compiles it from UT (Unidad de Transacciones, https://www.ut.com.sv/), whose "
        "servers do not answer outside El Salvador. Ember rounds to 0.01 TWh (10 GWh), so single months carry about "
        "+/-5 GWh of rounding.",
        "Updated weekly by GitHub Actions (el_salvador_gas.yml); the whole series is rebuilt each run (a few dozen rows).",
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--heat-rate", type=float, default=HEAT_RATE)
    args = ap.parse_args()

    r = requests.get(EMBER, timeout=300)
    r.raise_for_status()
    e = pd.read_csv(io.BytesIO(r.content), usecols=["Area", "Date", "Category", "Subcategory", "Variable", "Unit", "Value"])
    g = e[(e.Area == "El Salvador") & (e.Category == "Electricity generation") & (e.Variable == "Gas") & (e.Unit == "TWh")]
    if g.empty:
        print("Ember has no El Salvador gas generation rows.", flush=True)
        sys.exit(1)
    s = g.set_index(pd.to_datetime(g["Date"]))["Value"].sort_index() * 1000  # GWh
    s = s[s.index >= START]
    df = pd.DataFrame({"Gas_generation_GWh": s})
    df["Heat_rate_MMBtu_per_MWh"] = args.heat_rate
    df["Gas_use_MMBtu"] = df["Gas_generation_GWh"] * 1000 * args.heat_rate
    df["Gas_use_MMBtu_per_day"] = df["Gas_use_MMBtu"] / df.index.days_in_month
    df["Gas_use_mcm"] = df["Gas_use_MMBtu"] / MMBTU_PER_MCM
    df["LNG_equivalent_kt"] = df["Gas_use_MMBtu"] / MMBTU_PER_T_LNG / 1000
    df["Basis"] = "estimate: Ember gas generation x heat rate"
    df.index.name = "Month"
    df = df.round({"Gas_generation_GWh": 1, "Gas_use_MMBtu": 0, "Gas_use_MMBtu_per_day": 0, "Gas_use_mcm": 2,
                   "LNG_equivalent_kt": 2})
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    nts = notes(args.heat_rate)
    xlsx_notes.write_workbook(args.out, {"Gas use": df}, nts, {x for x in nts if x and x.isupper()})
    print(f"Saved {args.out}: {len(df)} months ({df.index.min():%b-%Y} to {df.index.max():%b-%Y})", flush=True)
    print(df.tail(12).to_string(), flush=True)


if __name__ == "__main__":
    main()
