"""
Chile natural gas demand by sector, monthly from 2021 (CLAUDE.md: South America gas demand by sector is 2021+).

No official Chilean gas consumption by sector exists for 2021 onwards: CNE's monthly 'Consumo_mensual_GN' file ends in
Oct 2018, the Anuario Estadistico de Energia has annual Tcal only (2021 edition last), and the Ministerio de
Energia / CNE Balance Nacional de Energia is annual (see discovery_archive/south_america/CHILE_GAS_DEMAND_DISCOVERY*.py).
So this workbook is DERIVED from raw series already in the repo (no network access; it only reads two workbooks):

  Power (ESTIMATE)       = SEN gas-fired generation (Gas_MWh, CNE 'Generacion Bruta' from Coordinador Electrico
                           Nacional data; chile_power_generation_daily.xlsx) x heat rate (assumption, default 7.5
                           MMBtu/MWh HHV - Chile's gas fleet mixes combined and open cycle).
  Petrochemical (Magallanes) = Argentine pipeline gas imported for methanol, CNE import workbook (customs data;
                           'Imports by use' sheet of chile_gas_imports.xlsx). Imports only: Methanex also uses domestic gas.
  Total supply           = gas imports (CNE) + domestic production (ENAP + CEOP, CNE/Ministerio de Energia).
                           Production ends Jun 2024 in the latest CNE file; later months have imports only, and no
                           residual is computed for them.
  Other demand (RESIDUAL)= total supply - power estimate - petrochemical. Industry, residential, commercial,
                           distributors' own use, stock change and the small export to Argentina cannot be separated.

Run order: after CHILE_POWER_DAILY.py and CHILE_GAS_IMPORTS.py (1st/15th, chile_gas_demand.yml).
Usage: python3 CHILE_GAS_DEMAND.py [--power PATH] [--imports PATH] [--out PATH] [--heat-rate 7.5]
"""
import argparse
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

D = "output/Data and Chart Outputs/"
MMBTU_PER_MCM = 36_374.0   # MMBtu per million m3 at 1,030 Btu/cf (gross), as add_charts.py uses
START = "2021-01-01"


def monthly(path, sheet, col):
    d = pd.read_excel(path, sheet_name=sheet)
    s = pd.Series(pd.to_numeric(d[col], errors="coerce").values, index=pd.to_datetime(d["Month"].astype(str), errors="coerce"))
    return s[s.index.notna()]


def build(power_path, imports_path, heat_rate):
    p = pd.read_excel(power_path, sheet_name="Daily", index_col=0)
    p.index = pd.to_datetime(p.index)
    g = p["Gas_MWh"].resample("MS").agg(["sum", "count"])
    full = g["sum"] / g["count"].where(g["count"] > 0) * g.index.days_in_month   # scale partial months
    out = pd.DataFrame(index=g.index)
    out["Gas_generation_GWh"] = (g["sum"] / 1000).round(1)
    out["Days_with_data"] = g["count"].astype(int)
    out["Power_est_mcm_per_day"] = (full * heat_rate / MMBTU_PER_MCM / g.index.days_in_month).round(3)
    for col, sheet, src in (("Imports_mcm_per_day", "Gas imports", "Imports_mcm_per_day_approx"),
                            ("Petrochemical_Magallanes_imports_mcm_per_day", "Imports by use",
                             "Pipeline_petrochemical_Magallanes_mcm_per_day"),
                            ("Domestic_production_mcm_per_day", "Domestic production", "Total_mcm_per_day")):
        try:
            out[col] = monthly(imports_path, sheet, src).reindex(out.index)
        except (ValueError, KeyError):   # sheet not in this workbook version
            out[col] = float("nan")
    out["Total_supply_mcm_per_day"] = (out["Imports_mcm_per_day"] + out["Domestic_production_mcm_per_day"]).round(3)
    out["Other_demand_residual_mcm_per_day"] = (
        out["Total_supply_mcm_per_day"] - out["Power_est_mcm_per_day"]
        - out["Petrochemical_Magallanes_imports_mcm_per_day"].fillna(0)).round(3)
    out["Heat_rate_MMBtu_per_MWh"] = heat_rate
    out = out[out.index >= START]
    out = out[out["Days_with_data"] > 0]
    out.index = out.index.strftime("%Y-%m")
    out.index.name = "Month"
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--power", default=D + "chile_power_generation_daily.xlsx")
    ap.add_argument("--imports", default=D + "chile_gas_imports.xlsx")
    ap.add_argument("--out", default=D + "chile_gas_demand_by_sector.xlsx")
    ap.add_argument("--heat-rate", type=float, default=7.5)
    a = ap.parse_args()
    out = build(a.power, a.imports, a.heat_rate)
    notes = [
        "UNITS",
        "Million m3 per day (mcm/d), monthly average, 1,030 Btu/cf (1 mcm = 36,374 MMBtu). Gas_generation_GWh = "
        "monthly SEN gas-fired generation.",
        "Power_est = ESTIMATE, not a published volume: gas-fired generation x heat rate "
        f"({a.heat_rate} MMBtu/MWh HHV, an assumption; error roughly +/-10-15%). Partial months scaled from the daily "
        "average. Other_demand_residual = total supply - power estimate - petrochemical: a DERIVED balance "
        "(industry, residential, commercial, own use, stock change and small exports together), empty from the month "
        "domestic production data end.",
        "",
        "COVERAGE",
        f"Chile, monthly, {out.index.min()} to {out.index.max()} (2021 on, per repo convention). Follows the power "
        "and imports workbooks; domestic production only to the last month of CNE's production file.",
        "",
        "SOURCE",
        "No official sector split exists from 2021: CNE Consumo_mensual_GN ends Oct 2018; the Anuario Estadistico de "
        "Energia and Balance Nacional de Energia are annual. Inputs: CNE Generacion Bruta (Coordinador Electrico "
        "Nacional data) via chile_power_generation_daily.xlsx; CNE import workbook and ENAP/CEOP production via "
        "chile_gas_imports.xlsx. https://www.cne.cl/estadisticas/",
        "Built by CHILE_GAS_DEMAND.py (chile_gas_demand.yml, 1st/15th, after the power and imports pulls).",
    ]
    xlsx_notes.write_workbook(a.out, {"Demand by sector": out}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {a.out}: {len(out)} months")
    print(out.tail(6).to_string())


if __name__ == "__main__":
    main()
