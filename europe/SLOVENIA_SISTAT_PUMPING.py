"""
Slovenia's electricity used for pumped storage (Avce), from the Statistical Office of Slovenia (SiStat table 1817602S, 'Electricity (GWh), Slovenia,
annually'), one workbook:

  output/Data and Chart Outputs/slovenia_sistat_pumping_daily.xlsx
    sheet "Daily":  date, PumpedStorageConsumption_MWh (annual pumping spread over the days in proportion to ENTSO-E's pumped-storage generation), Estimated
                    (1 = year not yet published by SiStat: last published pumping/generation ratio carried)
    sheet "Annual": SiStat measures per year, GWh (gross/net production, pumped-storage production, import, export, used for pumped storage,
                    losses in the network, final consumption)
    sheet "Units":  source and definitions

Why: ENTSO-E publishes Avce's pumped-storage generation (0.28-0.32 TWh a year) but NO pumping consumption, so the Europe supply/load balance counted the
output without the energy that went in (Slovenia's supply was 103.5-103.9% of load). SiStat's pumping figure (405 / 381 / 437 GWh in 2023 / 2024 / 2025; it is a constant
1.346 x the pumped-storage output every year, i.e. an efficiency assumption of 74.3% - the same as ENTSO-E's measured Kruonis and Coo ratios, 1.35-1.37 - not a meter reading) accounts for 0.38-0.44 of the 0.44-0.48 TWh gap. ENTSO-E's load excludes pumping by
definition. The Europe master subtracts this series from Slovenia's 'Pumped & battery (net)'.

Incremental: SiStat holds the whole annual history; the table is re-read each run (one request) and the daily series is rebuilt from the committed ENTSO-E
Slovenia workbook.
Usage: python3 SLOVENIA_SISTAT_PUMPING.py [--out-dir DIR]
"""
import argparse
import csv
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
FILE = "slovenia_sistat_pumping_daily.xlsx"
GEN_FILE = "slovenia_power_generation_daily.xlsx"
URL = "https://pxweb.stat.si/SiStatData/api/v1/en/Data/1817602S.px"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
KEEP = {"Gross production - Total": "Gross production", "Net production - Total": "Net production",
        "Gross production - Hydro - of which from pumped storage": "Pumped-storage production (gross)", "Import": "Import", "Export": "Export",
        "Used for pumped storage": "Used for pumped storage", "Losses in the network": "Losses in the network",
        "Final consumption - Total": "Final consumption"}


def annual():
    r = requests.post(URL, json={"query": [], "response": {"format": "csv"}}, headers=UA, timeout=(10, 90))
    r.raise_for_status()
    rows = list(csv.reader(io.StringIO(r.text)))
    years = [h.split()[0] for h in rows[0][1:]]   # '2025 (provisional data)' -> '2025'
    out = {}
    for row in rows[1:]:
        if row[0] in KEEP:
            out[KEEP[row[0]]] = {int(y): float(v) for y, v in zip(years, row[1:]) if v not in ("", "-", "...")}
    return pd.DataFrame(out).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    ann = annual()
    g = pd.read_excel(os.path.join(args.out_dir, GEN_FILE), sheet_name="Daily")
    g["date"] = pd.to_datetime(g["date"])
    gen = pd.to_numeric(g.set_index("date")["PumpedStorage_MWh"], errors="coerce").dropna()
    ratio = (ann["Used for pumped storage"] / ann["Pumped-storage production (gross)"]).dropna()   # GWh in per GWh out
    print("pumping / pumped-storage output by year:\n" + ratio.round(3).to_string(), flush=True)
    last = ratio.index.max()
    daily = pd.DataFrame(index=gen.index)
    daily["year"] = daily.index.year
    daily["Estimated"] = (~daily["year"].isin(ratio.index)).astype(int)
    daily["PumpedStorageConsumption_MWh"] = (gen * daily["year"].map(lambda y: ratio.get(y, ratio[last]))).round(1)
    daily = daily[["PumpedStorageConsumption_MWh", "Estimated"]]
    daily.index.name = "date"
    ann.index.name = "year"
    cmp = pd.DataFrame({"SiStat GWh": ann["Used for pumped storage"], "spread over ENTSO-E days GWh": daily["PumpedStorageConsumption_MWh"].groupby(daily.index.year).sum() / 1000}).dropna()
    print(cmp.round(1).to_string(), flush=True)
    lines = ["Slovenia - electricity used for pumped storage (Statistical Office of Slovenia, SiStat)", "",
             "Source", "Statistical Office of the Republic of Slovenia (SURS), SiStat table 1817602S 'Electricity (GWh), Slovenia, annually', https://pxweb.stat.si/SiStatData/pxweb/en/Data/-- . Free, no key.",
             "", "Units and definitions",
             "Daily: MWh of pumping consumption per UTC day = SiStat's annual 'Used for pumped storage' x (ENTSO-E pumped-storage generation of the day / the year's total), "
             "i.e. the annual statistic given a daily shape from ENTSO-E (Avce pumps and generates in the same daily cycle). Estimated = 1 where SiStat has not yet published "
             "the year (the latest published pumping per unit of pumped-storage output, 1.35, is carried). Annual: SiStat measures in GWh (the latest year is provisional). "
             "ENTSO-E publishes Avce's generation but not its pumping consumption, and its load excludes pumping, so a supply/load balance needs this series.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; daily {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}; SiStat years {ann.index.min()}-{ann.index.max()}"]
    xlsx_notes.write_workbook(os.path.join(args.out_dir, FILE), {"Daily": daily, "Annual": ann}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
