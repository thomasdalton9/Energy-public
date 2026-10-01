"""
Panama monthly natural-gas (LNG) use for power - ESTIMATE from CND gas-fired
generation and a stated heat rate.

Panama's only gas consumers of scale are the LNG-fired power plants on the
Atlantic coast at Colon: AES Colon (381 MW combined cycle, with the Costa Norte
LNG terminal, since 2018) and Gatun (Sinolam, ~670 MW combined cycle, from 2024).
Neither CND, the Secretaria de Energia nor ASEP publishes gas volumes burned or
LNG imported in a machine-readable form (CND's monthly operations report gives
fuel prices only; see discovery_archive/south_america/CENTRAL_AMERICA_POWER_DISCOVERY6.py),
so gas use is estimated:

    gas use (MMBtu, HHV) = CND 'Gas Natural' energy delivered (MWh) x heat rate

with the heat rate a stated assumption (default 7.0 MMBtu/MWh HHV, i.e.
7,000 Btu/kWh, ~49% LHV efficiency - typical for modern F/H-class combined
cycles running at part load; override with --heat-rate). Conversions:
1 million m3 of gas = 36,300 MMBtu (HHV 38.3 MJ/m3); 1 tonne LNG = 52.0 MMBtu.

Input: the daily CND workbook written by PANAMA_CND_GENERATION_DAILY.py
(sheet 'Daily', Gas_MWh). No network access; runs right after that pull.

Usage: python3 PANAMA_GAS.py [--power PATH] [--out PATH] [--heat-rate 7.0]
"""

print("STARTING", flush=True)

import argparse
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

POWER = "output/Data and Chart Outputs/panama_power_generation_daily.xlsx"
OUT = "output/Data and Chart Outputs/panama_gas.xlsx"
MMBTU_PER_MCM = 36_300.0   # MMBtu (HHV) per million m3
MMBTU_PER_T_LNG = 52.0     # MMBtu (HHV) per tonne of LNG
MMBTU_PER_MMCF = 1_037.0   # MMBtu (HHV) per million cubic feet


def build(power_path, heat_rate):
    d = pd.read_excel(power_path, sheet_name="Daily", index_col=0)
    d.index = pd.to_datetime(d.index)
    g = d["Gas_MWh"].resample("MS").agg(["sum", "count"])
    days_in_month = g.index.days_in_month
    out = pd.DataFrame(index=g.index)
    out.index.name = "Month"
    out["Gas_generation_MWh"] = g["sum"].round(1)
    out["Days_with_data"] = g["count"].astype(int)
    out["Days_in_month"] = days_in_month
    # scale partial months (missing days) to the whole month from the daily average
    full = out["Gas_generation_MWh"] / out["Days_with_data"].where(out["Days_with_data"] > 0) * out["Days_in_month"]
    out["Heat_rate_MMBtu_per_MWh"] = heat_rate
    out["Gas_use_MMBtu_est"] = (full * heat_rate).round(0)
    out["Gas_use_MMBtu_per_day_est"] = (out["Gas_use_MMBtu_est"] / out["Days_in_month"]).round(0)
    out["Gas_use_MMcf_per_day_est"] = (out["Gas_use_MMBtu_per_day_est"] / MMBTU_PER_MMCF).round(1)
    out["Gas_use_mcm_est"] = (out["Gas_use_MMBtu_est"] / MMBTU_PER_MCM).round(2)
    out["LNG_equivalent_kt_est"] = (out["Gas_use_MMBtu_est"] / MMBTU_PER_T_LNG / 1000).round(1)
    out["Complete_month"] = out["Days_with_data"] >= out["Days_in_month"] - 1
    return out[out["Days_with_data"] > 0]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--power", default=POWER)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--heat-rate", type=float, default=7.0, help="MMBtu (HHV) per MWh delivered")
    args = ap.parse_args()
    out = build(args.power, args.heat_rate)
    notes = [
        "UNITS",
        "ESTIMATE - not a published volume. Gas_use_MMBtu_est = CND gas-fired energy (MWh) x heat rate "
        f"({args.heat_rate} MMBtu/MWh, HHV). Months with missing days are scaled up from the daily average "
        "(Days_with_data / Days_in_month shown). MMBtu/day, MMcf/day, million m3 (mcm) and kt of LNG equivalent.",
        f"Conversions: 1 mcm = {MMBTU_PER_MCM:,.0f} MMBtu; 1 MMcf = {MMBTU_PER_MMCF:,.0f} MMBtu; "
        f"1 t LNG = {MMBTU_PER_T_LNG} MMBtu (all HHV).",
        "",
        "COVERAGE",
        f"{out.index.min():%b %Y} to {out.index.max():%b %Y} (follows panama_power_generation_daily.xlsx).",
        "",
        "SOURCE",
        "Gas-fired generation: CND / ETESA Panama daily report ('Reporte Diario', ENTREGADO AL SISTEMA - Gas Natural), "
        "https://www.cnd.com.pa/index.php/informes/categoria/informes-de-operaciones, via "
        "panama_power_generation_daily.xlsx (PANAMA_CND_GENERATION_DAILY.py).",
        "No official gas-volume series: CND's monthly operations report gives fuel prices only; the Secretaria de "
        "Energia (https://www.energia.gob.pa/) and ASEP (https://www.asep.gob.pa/) were searched for machine-readable "
        "LNG import / gas consumption statistics (CENTRAL_AMERICA_POWER_DISCOVERY6.py). The only gas figure found, "
        "the Secretaria's annual energy balance (BALANCES-DE-ENERGIA-1970-2025.xls), shows natural-gas input to "
        "power of ~5.6 thousand boe in 2024, hundreds of times too small for CND's ~2.2 TWh of gas-fired output, "
        "so it is not used.",
        "",
        "METHOD",
        "Plants: AES Colon (381 MW combined cycle, Costa Norte LNG terminal, Colon, since 2018) and Gatun "
        "(Sinolam, ~670 MW combined cycle, from 2024). Both burn regasified LNG from Costa Norte.",
        f"Heat rate {args.heat_rate} MMBtu/MWh (7,000 Btu/kWh HHV, ~49% LHV efficiency) is an assumption for "
        "modern combined cycles at part load; actual heat rates are not published. Error of the estimate is "
        "roughly +/-10% from the heat rate alone.",
        "Updated daily by GitHub Actions (panama_power_generation.yml) after the CND pull.",
    ]
    titles = {x for x in notes if x and x.isupper()}
    xlsx_notes.write_workbook(args.out, {"Gas use": out}, notes, titles)
    print(f"Saved {args.out}: {len(out)} months", flush=True)
    print(out.tail(6).to_string(), flush=True)


if __name__ == "__main__":
    main()
