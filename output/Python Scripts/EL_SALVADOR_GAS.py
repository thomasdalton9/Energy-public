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
  - DGEHM's hydrocarbon statistics (sinapp.dgehm.gob.sv) also do not answer
    outside El Salvador (PROBE15).
So gas use is estimated from monthly gas-fired generation x a stated heat
rate:
  gas generation : SIGET monthly NET generation, technology 'GNL', from
                   el_salvador_power_generation_daily.xlsx (EL_SALVADOR_SIGET_POWER.py)
                   where it has the month (Jan-2023 to Dec-2025 as of Oct-2026);
                   other months from Ember monthly electricity data, El
                   Salvador, 'Gas' (compiled from UT; rounded to 10 GWh);
  heat rate      : 8.2 MMBtu (HHV) per MWh, i.e. ~46% net LHV efficiency,
                   typical for Wartsila 50DF engines on gas at part load.
Conversions: 1 mcm of regasified LNG ~ 36,300 MMBtu (HHV 38.3 MJ/m3);
1 tonne LNG ~ 51.7 MMBtu (HHV).
Check: 2023 estimate 2,481 GWh x 8.2 = 20.3 TBtu HHV = 21,460 TJ HHV, about
19,370 TJ on a net (LHV) basis, vs 18,749 TJ of natural-gas imports in
DGEHM's 2023 energy balance (+3%).

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
SIGET_XLSX = "output/Data and Chart Outputs/el_salvador_power_generation_daily.xlsx"
SIGET_PAGE = ("https://www.siget.gob.sv/gerencias/electricidad/informe-de-mercado-y-estadisticas-electricas/"
              "estadisticas-electricas-bi/")


def siget_gas(path):
    """SIGET monthly net LNG generation (GWh) from the power workbook, or an empty series."""
    try:
        d = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.Series(dtype=float)
    d.index = pd.to_datetime(d.index, errors="coerce")
    if "Gas_MWh" not in d:
        return pd.Series(dtype=float)
    return (d["Gas_MWh"].dropna() / 1000).sort_index()


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
        "Check: the 2023 estimate (2,481 GWh x 8.2 = 21,460 TJ HHV, about 19,370 TJ LHV) is 3% above the 18,749 TJ of "
        "natural-gas imports in DGEHM's 2023 national energy balance (https://estadisticas.dgehm.gob.sv/).",
        "",
        "SOURCE",
        "Gas generation (column Gas_generation_source says which, per month): (1) SIGET monthly NET generation, "
        "technology 'GNL', from el_salvador_power_generation_daily.xlsx (SIGET 'Visualizador dinamico de Estadisticas "
        f"Electricas', {SIGET_PAGE}), used for every month it holds; (2) otherwise Ember monthly electricity data "
        f"(CC-BY-4.0), El Salvador, 'Gas', {EMBER}, which Ember compiles from UT (https://www.ut.com.sv/, whose "
        "servers do not answer outside El Salvador) and rounds to 10 GWh. The two agree within 0.3% a year in "
        "2023-2025.",
        "Updated weekly by GitHub Actions (el_salvador_gas.yml); the whole series is rebuilt each run (a few dozen rows).",
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--heat-rate", type=float, default=HEAT_RATE)
    ap.add_argument("--siget", default=SIGET_XLSX, help="SIGET power workbook (EL_SALVADOR_SIGET_POWER.py output)")
    args = ap.parse_args()

    r = requests.get(EMBER, timeout=300)
    r.raise_for_status()
    e = pd.read_csv(io.BytesIO(r.content), usecols=["Area", "Date", "Category", "Subcategory", "Variable", "Unit", "Value"])
    g = e[(e.Area == "El Salvador") & (e.Category == "Electricity generation") & (e.Variable == "Gas") & (e.Unit == "TWh")]
    if g.empty:
        print("Ember has no El Salvador gas generation rows.", flush=True)
        sys.exit(1)
    s = g.set_index(pd.to_datetime(g["Date"]))["Value"].sort_index() * 1000  # GWh
    sg = siget_gas(args.siget)
    src = pd.Series("Ember (from UT)", index=s.index)
    s = pd.concat([s[~s.index.isin(sg.index)], sg]).sort_index()
    src = src.reindex(s.index).fillna("Ember (from UT)")
    src[src.index.isin(sg.index)] = "SIGET (net, GNL)"
    print(f"Gas generation: {len(sg)} months from SIGET, the rest from Ember", flush=True)
    s, src = s[s.index >= START], src[src.index >= START]
    df = pd.DataFrame({"Gas_generation_GWh": s, "Gas_generation_source": src})
    df["Heat_rate_MMBtu_per_MWh"] = args.heat_rate
    df["Gas_use_MMBtu"] = df["Gas_generation_GWh"] * 1000 * args.heat_rate
    df["Gas_use_MMBtu_per_day"] = df["Gas_use_MMBtu"] / df.index.days_in_month
    df["Gas_use_mcm"] = df["Gas_use_MMBtu"] / MMBTU_PER_MCM
    df["LNG_equivalent_kt"] = df["Gas_use_MMBtu"] / MMBTU_PER_T_LNG / 1000
    df["Basis"] = "estimate: gas generation x heat rate"
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
