"""
Switzerland daily electricity production by type and national consumption from the Swiss Federal Office of Energy's open data
(energiedashboard.ch, Swissgrid data), one workbook in the same daily layout as the ENTSO-E country workbooks:

  output/Data and Chart Outputs/switzerland_swissgrid_power_daily.xlsx
    sheet "Daily": date, <Fuel>_MWh (Hydro = run-of-river + storage, Nuclear, Wind, Solar, Other = thermal incl. waste/biomass, Gas = 0),
                   Total_MWh, Load_MWh (= Landesverbrauch, national consumption)
    sheet "By type (GWh)": the BFE categories as published (Flusskraft, Speicherkraft, Kernkraft, ...), GWh per day
    sheet "Units": source and definitions

Why: ENTSO-E's Swiss generation is incomplete and its coverage changes by year (non-pumped hydro 10 TWh in 2022, 24 TWh in 2025), which
made the Swiss power balance 30% short in 2023-24. Swissgrid's own statistics cover the whole country.

Sources (free, no key; the server rejects requests without a browser-like User-Agent)
  https://www.uvek-gis.admin.ch/BFE/ogd/104/ogd104_stromproduktion_swissgrid.csv   Datum, Energietraeger, Produktion_GWh (from 2015)
  https://www.uvek-gis.admin.ch/BFE/ogd/103/ogd103_stromverbrauch_swissgrid_lv_und_endv.csv   Datum, Landesverbrauch_GWh, Endverbrauch_GWh

Both files are small whole-file CSVs that hold the full history, so each run re-downloads and rewrites them. Recent days are restated by Swissgrid.

Usage: python3 SWITZERLAND_SWISSGRID_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import io
import os
import sys
from datetime import datetime, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "switzerland_swissgrid_power_daily.xlsx"
PROD_URL = "https://www.uvek-gis.admin.ch/BFE/ogd/104/ogd104_stromproduktion_swissgrid.csv"
BAL_URL = "https://www.bfe-ogd.ch/ogd35/ogd35_schweizerische_elektrizitaetsbilanz_monatswerte.csv"
CONS_URL = "https://www.uvek-gis.admin.ch/BFE/ogd/103/ogd103_stromverbrauch_swissgrid_lv_und_endv.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
           "Accept": "text/csv,application/json,text/html,*/*;q=0.8", "Accept-Language": "en-GB,en;q=0.9", "Referer": "https://www.bfe.admin.ch/"}


def group(name):
    n = str(name).lower()
    if "fluss" in n or "speicher" in n or "wasser" in n or "hydro" in n:
        return "Hydro"
    if "kern" in n:
        return "Nuclear"
    if "wind" in n:
        return "Wind"
    if "photo" in n or "solar" in n or "pv" in n:
        return "Solar"
    return "Other"        # thermal power stations (waste, biomass, fossil) and anything unrecognised


def read_csv(url):
    r = requests.get(url, headers=HEADERS, timeout=120)
    r.raise_for_status()
    if r.text.lstrip().startswith("<"):
        raise RuntimeError(f"{url}: HTML answer ({r.text[:120]!r})")
    return pd.read_csv(io.StringIO(r.text))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    prod = read_csv(PROD_URL)
    cons = read_csv(CONS_URL)
    prod["Datum"] = pd.to_datetime(prod["Datum"], errors="coerce")
    cons["Datum"] = pd.to_datetime(cons["Datum"], errors="coerce")
    prod = prod.dropna(subset=["Datum"])
    cons = cons.dropna(subset=["Datum"]).drop_duplicates("Datum", keep="last").set_index("Datum").sort_index()   # the file repeats some dates
    cats = sorted(prod["Energietraeger"].dropna().unique())
    print("categories:", {c: group(c) for c in cats}, flush=True)
    by_type = prod.pivot_table(index="Datum", columns="Energietraeger", values="Produktion_GWh", aggfunc="sum").sort_index()
    by_type = by_type[by_type.index >= args.start]
    daily = pd.DataFrame(index=by_type.index)
    for g in ("Hydro", "Nuclear", "Wind", "Solar", "Other"):
        cols = [c for c in by_type.columns if group(c) == g]
        daily[f"{g}_MWh"] = by_type[cols].sum(axis=1, min_count=1) * 1000.0 if cols else 0.0
    daily["Gas_MWh"] = 0.0
    for c in ("PumpedStorage_MWh", "Coal_MWh", "Oil_MWh", "Bioenergy_MWh", "Storage_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh"):
        daily[c] = 0.0
    # Pumping: Swissgrid's Speicherkraft is gross of pumped-storage output, and the electricity used for pumping is part of what the grid
    # supplies. BFE's monthly Swiss electricity balance (ogd35) has the pumping consumption and the physical imports/exports; they are spread
    # evenly over the days of each month (a monthly figure, labelled as such on the Units sheet).
    pump_note = "Pumping consumption and net imports not available (BFE ogd35 download failed)"
    monthly = None
    try:
        monthly = read_csv(BAL_URL)
        monthly["month"] = pd.to_datetime(dict(year=monthly["Jahr"], month=monthly["Monat"], day=1))
        monthly = monthly.set_index("month").sort_index()
        dim = pd.Series(daily.index.days_in_month, index=daily.index)
        mkey = daily.index.to_period("M").to_timestamp()
        for src, dst, sign in (("Verbrauch_Speicherpumpen_GWh", "PumpedStorageConsumption_MWh", 1.0), ("Einfuhr_GWh", "Imports_MWh", 1.0), ("Ausfuhr_GWh", "Exports_MWh", 1.0)):
            daily[dst] = (monthly[src].reindex(mkey).to_numpy() * 1000.0 / dim.to_numpy()) * sign
        daily["NetImports_MWh"] = daily["Imports_MWh"] - daily["Exports_MWh"]
        daily = daily.drop(columns=["Imports_MWh", "Exports_MWh"])
        pump_note = ("PumpedStorageConsumption_MWh and NetImports_MWh (physical imports less exports) come from BFE's monthly Swiss electricity "
                     "balance (ogd35), spread evenly over the days of each month")
    except Exception as ex:  # noqa: BLE001
        print("ogd35 failed:", ex, flush=True)
    print(pump_note, flush=True)
    daily["Total_MWh"] = daily[["Hydro_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh", "Other_MWh"]].sum(axis=1)
    daily["Load_MWh"] = (cons["Landesverbrauch_GWh"] * 1000.0).reindex(daily.index)
    daily = daily.dropna(subset=["Total_MWh"]).round(1)
    daily.index.name = "date"
    by_type = by_type.round(3)
    by_type.index.name = "date"
    ann = (daily.groupby(daily.index.year).sum() / 1e6).round(1)
    print("TWh per year:\n" + ann[["Hydro_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh", "Other_MWh", "Total_MWh", "Load_MWh"]].T.to_string(), flush=True)
    print("days per year:", daily.groupby(daily.index.year).size().to_dict())
    lines = ["Switzerland - daily electricity production by type and national consumption (Swiss Federal Office of Energy / Swissgrid)", "",
             "Source", "BFE open data (energiedashboard.ch), Swissgrid statistics: ogd104 production by energy carrier, ogd103 national consumption. "
             "https://www.energiedashboard.ch . Free, no key.",
             "", "Units and definitions",
             "MWh per day (the source is GWh). Hydro = run-of-river (Flusskraft) + storage (Speicherkraft); Nuclear (Kernkraft); Wind; "
             "Solar = photovoltaics; Other = thermal power stations (incl. waste, biomass); Load_MWh = Landesverbrauch (national consumption, "
             "including grid losses). The 'By type (GWh)' sheet keeps the categories as published. ENTSO-E's Swiss generation is incomplete "
             "(its hydro rose from 10 to 24 TWh between 2022 and 2025 as reporting widened), so this feed replaces it in the Europe balances.",
             pump_note + ".",
             "Whole-file download each run (the files hold the full history); recent days are restated.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(daily)} days, {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": daily, "By type (GWh)": by_type}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
