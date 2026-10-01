"""
Ecuador daily power generation by type, in the standard layout every
country's raw grid-operator workbook uses (sheet 'Daily': date, Hydro_MWh,
Gas_MWh, ..., Total_MWh), from CENACE.

Reads south_america/ec_generation_daily.csv, which ECUADOR_CENACE.py
(same workflow, .github/workflows/ecuador_daily.yml) upserts from CENACE's
Informacion Operativa page every run. That page only ever shows the last
complete day and CENACE keeps no public archive (checked in
discovery_archive/south_america/PERU_ECUADOR_POWER_DISCOVERY2-4.py: the
info-operativa folder is 403, the WordPress media library has no
operation files, and the Wayback Machine holds only ~30 scattered copies
since 2020) - so history starts at the first run, 2026-09-26, and grows
by one day per day. It cannot go back to 2021.

Usage: python3 ECUADOR_POWER_DAILY.py [--csv PATH] [--out PATH]
"""

print("STARTING", flush=True)

import argparse
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

DEFAULT_CSV = os.path.join(ROOT, "south_america", "ec_generation_daily.csv")
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "ecuador_power_generation_daily.xlsx")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", default=DEFAULT_CSV)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    raw = pd.read_csv(args.csv, index_col=0, parse_dates=True).sort_index()
    raw = raw[~raw.index.duplicated(keep="last")]
    col = lambda c: raw[c] if c in raw else pd.Series(float("nan"), index=raw.index)  # noqa: E731
    thermal = col("thermal_mwh")
    gas = col("thermal_gas_mwh")
    oil = col("thermal_oil_mwh").fillna(thermal - gas.fillna(0))

    daily = pd.DataFrame(index=raw.index.date)
    daily.index.name = "date"
    daily["Hydro_MWh"] = col("hydro_mwh").values
    daily["Gas_MWh"] = gas.values
    daily["Oil_MWh"] = oil.values
    daily["Other_MWh"] = col("renewable_mwh").values
    daily["Total_MWh"] = (col("hydro_mwh") + thermal + col("renewable_mwh").fillna(0)).values
    daily = daily.dropna(subset=["Hydro_MWh", "Total_MWh"], how="all").round(1)

    detail = raw[[c for c in ["total_mwh", "hydro_mwh", "thermal_mwh", "thermal_gas_mwh", "thermal_oil_mwh",
                              "renewable_mwh", "imports_mwh", "exports_mwh"] if c in raw]].copy()
    detail.index = detail.index.date
    detail.index.name = "date"

    notes = [
        "UNITS",
        "'Daily': energy generated per day in MWh by type (CENACE's own daily MWh, preliminary SCADA data). "
        "Total_MWh = domestic generation = hydro + thermal + non-conventional renewables (imports excluded).",
        "'CENACE raw': the same days as CENACE reports them, incl. imports/exports and CENACE's 'Produccion total' "
        "(which includes imports).",
        "",
        "CATEGORY MAPPING",
        "Hydro_MWh = CENACE 'Hidraulica'",
        "Gas_MWh = CENACE thermal 'Gas Natural' (Machala plants on Amistad-field gas)",
        "Oil_MWh = CENACE thermal 'Termica' (fuel oil, diesel, residual, crude - the rest of thermal); charted as "
        "'Other Fossil'",
        "Other_MWh = CENACE 'R. No Convencional' (non-conventional renewables: wind, solar, biomass/bagasse, "
        "biogas - CENACE doesn't split them); charted as 'Other Renewables'",
        "No coal, nuclear, or separately reported wind/solar/bioenergy in CENACE's daily report, so those columns "
        "are omitted.",
        "",
        "COVERAGE",
        f"{daily.index.min()} to {daily.index.max()} ({len(daily)} days). HISTORY IS LIMITED: CENACE publishes only "
        "the last complete day and keeps no public archive, so the series starts at the first run of "
        "ECUADOR_CENACE.py (2026-09-26) and cannot be backfilled to 2021. Days the page wasn't read are gaps.",
        "",
        "SOURCE",
        "CENACE Informacion Operativa (https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm), "
        "'Informacion operativa diaria' tab, read twice a day by south_america/ECUADOR_CENACE.py into "
        "south_america/ec_generation_daily.csv; this workbook is rebuilt from that file by "
        "south_america/ECUADOR_POWER_DAILY.py (.github/workflows/ecuador_daily.yml).",
    ]
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "CENACE raw": detail}, notes,
                              {"UNITS", "CATEGORY MAPPING", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: {len(daily)} day(s) {daily.index.min()} .. {daily.index.max()}", flush=True)
    print(daily.tail().to_string(), flush=True)


if __name__ == "__main__":
    main()
