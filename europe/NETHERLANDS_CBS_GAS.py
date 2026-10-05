"""
Netherlands natural gas balance from CBS StatLine 86103NED 'Aardgasbalans; aanbod en verbruik' (monthly), as a workbook:

  output/Data and Chart Outputs/netherlands_cbs_gas_monthly.xlsx
    sheet "Monthly":  month, <item>_GWh for the main balance items (CBS mln m3 of Groningen-equivalent gas, 35.17 MJ/m3 gross, x 9.769 kWh/m3):
                      Production (winning uit de bodem), Biomethane (productie uit andere bronnen), Imports_pipeline (+ via Norway/Germany/Belgium/UK/Denmark),
                      Imports_LNG, Exports_pipeline (+ to Germany/Belgium/UK), Exports_LNG, Bunkering, StockChange, Consumption (totaal verbruik)
    sheet "Units":    source and definitions

Why: ENTSOG's Dutch production entries run 15-17 TWh a year above the national statistic (2023: 126 vs CBS 109.5) and its distribution + final-consumer
exits are about 6 TWh below CBS total consumption, which left the combined Germany + Netherlands gas balance about 4% long. CBS (Statistics
Netherlands, from the operators' and NLOG reporting) is a national statistics office, an accepted raw source when used as a labelled replacement feed.
The table is small and re-read whole each run (figures are revised).

Usage: python3 NETHERLANDS_CBS_GAS.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
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
FILE = "netherlands_cbs_gas_monthly.xlsx"
URL = "https://opendata.cbs.nl/ODataApi/odata/86103NED/TypedDataSet"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept": "application/json"}
KWH_PER_M3 = 35.17 / 3.6       # Groningen-equivalent gas, gross calorific value; mln m3 x kWh/m3 = GWh
MAP = {"WinningUitDeBodem_2": "Production_GWh", "ProductieUitAndereBronnen_3": "Biomethane_GWh",
       "InvoerVanGasvormigAardgasTotaal_4": "Imports_pipeline_GWh", "InvoerViaNoorwegen_5": "Imports_via_Norway_GWh",
       "InvoerViaDuitsland_6": "Imports_via_Germany_GWh", "InvoerViaBelgie_7": "Imports_via_Belgium_GWh",
       "InvoerViaVerenigdKoninkrijk_8": "Imports_via_UK_GWh", "InvoerViaDenemarken_9": "Imports_via_Denmark_GWh",
       "InvoerVloeibaarAardgasLNGTotaal_10": "Imports_LNG_GWh", "UitvoerVanGasvormigAardgasTotaal_18": "Exports_pipeline_GWh",
       "UitvoerNaarDuitsland_19": "Exports_to_Germany_GWh", "UitvoerNaarBelgie_20": "Exports_to_Belgium_GWh",
       "UitvoerNaarVerenigdKoninkrijk_21": "Exports_to_UK_GWh", "UitvoerVanVloeibaarAardgasLNG_22": "Exports_LNG_GWh",
       "Bunkering_23": "Bunkering_GWh", "Voorraadmutatie_24": "StockChange_GWh", "TotaalVerbruik_25": "Consumption_GWh"}


def fetch():
    rows, url = [], URL + "?$format=json&$top=2000"
    while url:
        j = requests.get(url, headers=H, timeout=120).json()
        rows += j.get("value", [])
        url = j.get("odata.nextLink")
    df = pd.DataFrame(rows)
    df = df[df["Perioden"].str.contains("MM")].copy()
    df["month"] = pd.to_datetime(df["Perioden"].str.replace("MM", "-", regex=False) + "-01", format="%Y-%m-%d")
    return df.set_index("month").sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    raw = fetch()
    raw = raw[raw.index >= args.start]
    mon = pd.DataFrame({dst: pd.to_numeric(raw[src], errors="coerce") * KWH_PER_M3 for src, dst in MAP.items() if src in raw})
    mon = mon[mon["Production_GWh"].notna() & mon["Consumption_GWh"].notna()].round(1)
    mon.index.name = "month"
    ann = (mon.groupby(mon.index.year).sum() / 1000).round(1)
    print("TWh per year:\n" + ann.T.to_string(), flush=True)
    lines = ["Netherlands - monthly natural gas balance (CBS StatLine 86103NED)", "",
             "Source", "Statistics Netherlands (CBS) 'Aardgasbalans; aanbod en verbruik', https://opendata.cbs.nl/ODataApi/odata/86103NED . Free, no key.",
             "", "Units and definitions",
             "GWh per month = CBS mln m3 (Groningen-equivalent gas, 35.17 MJ/m3 gross calorific value) x 9.769 kWh/m3. Production = winning uit de bodem "
             "(gas from Dutch fields); Biomethane = productie uit andere bronnen (chiefly upgraded biogas); Consumption = totaal verbruik (energy companies "
             "incl. power plants, own use at production and transport, flaring, final consumers). Imports/exports are pipeline gas by partner and LNG. "
             "StockChange (Voorraadmutatie): positive = stock WITHDRAWAL (a supply item: production + imports + StockChange - exports - bunkering = consumption closes only with this sign; it matches the AGSI+ Dutch net withdrawal). The Europe master gives the monthly production and consumption a daily shape from ENTSOG's Dutch entries/exits and keeps the CBS monthly totals exactly. CBS figures are revised; the table is re-read whole each run.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {mon.index.min():%Y-%m} to {mon.index.max():%Y-%m}"]
    xlsx_notes.write_workbook(os.path.join(args.out_dir, FILE), {"Monthly": mon}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
