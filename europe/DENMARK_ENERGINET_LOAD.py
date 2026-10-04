"""
Denmark electricity consumption and production settlement from Energinet's Energi Data Service (dataset ProductionConsumptionSettlement,
hourly, DK1 + DK2 summed to UTC days), one workbook:

  output/Data and Chart Outputs/denmark_energinet_load_daily.xlsx
    sheet "Daily": date, Load_MWh (gross consumption incl. grid losses and power-to-heat), PowerToHeat_MWh, Production_MWh (central + local +
                   commercial + self-consumption + wind + solar + hydro + unknown), NetImports_MWh (sum of the exchange columns), GridLoss_MWh
    sheet "Units": source and definitions

Why: ENTSO-E's Danish 'actual total load' is 4-7% below the generation + net imports it is metered against (supply/load 1.04-1.07 for 2022-25): Energinet's
settlement gross consumption includes electric boilers/heat pumps (power-to-heat, 2.7 TWh in 2025) and grid losses, and generation + exchanges equal it
by construction. The Europe master uses this Load for Denmark; generation by fuel and flows stay ENTSO-E (flows agree with Energinet's exchanges).

Incremental: the committed workbook is the history store; each run re-reads the last 45 days (settlement is revised) and any later days.
Usage: python3 DENMARK_ENERGINET_LOAD.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import os
import sys
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "denmark_energinet_load_daily.xlsx"
URL = "https://api.energidataservice.dk/dataset/ProductionConsumptionSettlement"
REVISION_DAYS = 45
PROD = ["CentralPowerMWh", "LocalPowerMWh", "CommercialPowerMWh", "LocalPowerSelfConMWh", "OffshoreWindLt100MW_MWh", "OffshoreWindGe100MW_MWh",
        "OnshoreWindLt50kW_MWh", "OnshoreWindGe50kW_MWh", "HydroPowerMWh", "SolarPowerLt10kW_MWh", "SolarPowerGe10Lt40kW_MWh",
        "SolarPowerGe40kW_MWh", "SolarPowerSelfConMWh", "UnknownProdMWh"]
EXCH = ["ExchangeNO_MWh", "ExchangeSE_MWh", "ExchangeGE_MWh", "ExchangeNL_MWh", "ExchangeGB_MWh", "ExchangeGreatBelt_MWh"]
LOSS = ["GridLossTransmissionMWh", "GridLossInterconnectorsMWh", "GridLossDistributionMWh"]


def fetch(d0, d1):
    out = []
    t = d0
    while t < d1:   # one year per request keeps the response small
        t2 = min(d1, t + timedelta(days=366))
        r = requests.get(URL, params={"start": f"{t:%Y-%m-%dT00:00}", "end": f"{t2:%Y-%m-%dT00:00}", "limit": 0}, timeout=300)
        r.raise_for_status()
        out += r.json().get("records", [])
        t = t2
    d = pd.DataFrame(out)
    if d.empty:
        return d
    d["day"] = pd.to_datetime(d["HourUTC"]).dt.floor("D")
    g = d.groupby("day")
    n = g["HourUTC"].nunique()
    daily = pd.DataFrame({
        "Load_MWh": g["GrossConsumptionMWh"].sum(min_count=1),
        "PowerToHeat_MWh": g["PowerToHeatMWh"].sum(min_count=1),
        "Production_MWh": d[PROD].fillna(0).groupby(d["day"]).sum().sum(axis=1),
        "NetImports_MWh": d[EXCH].fillna(0).groupby(d["day"]).sum().sum(axis=1),
        "GridLoss_MWh": d[LOSS].fillna(0).groupby(d["day"]).sum().sum(axis=1)})
    daily = daily[n.reindex(daily.index) >= 24]   # complete days only (24 hours, two price areas share the hour)
    daily.index.name = "date"
    return daily.dropna(subset=["Load_MWh"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    old = pd.DataFrame()
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.set_index("date").sort_index()
        except Exception:  # noqa: BLE001
            old = pd.DataFrame()
    start = datetime.strptime(args.start, "%Y-%m-%d")
    if len(old):
        start = max(start, old.index.max().to_pydatetime() - timedelta(days=REVISION_DAYS))
    end = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(days=1)
    new = fetch(start, end)
    print(f"fetched {len(new)} days from {start:%Y-%m-%d}", flush=True)
    daily = pd.concat([old[~old.index.isin(new.index)], new]).sort_index() if len(old) else new
    if daily.empty:
        print("no data")
        return
    daily = daily.round(1)
    ann = (daily.groupby(daily.index.year).sum() / 1e6).round(2)
    print("TWh per year:\n" + ann.T.to_string(), flush=True)
    lines = ["Denmark - electricity consumption and production settlement (Energinet, DK1 + DK2)", "",
             "Source", "Energinet, Energi Data Service dataset ProductionConsumptionSettlement, https://www.energidataservice.dk/tso-electricity/ProductionConsumptionSettlement . Free, no key.",
             "", "Units and definitions",
             "MWh per UTC day, hourly values summed over DK1 and DK2; a day is kept only when it has 24 hours. Load_MWh = GrossConsumptionMWh (consumption "
             "incl. grid losses and power-to-heat); PowerToHeat_MWh = electric boilers and heat pumps (part of Load); Production_MWh = central + local + "
             "commercial + self-consumed production + wind + solar + hydro + unknown; NetImports_MWh = sum of the exchange columns (positive = import); "
             "GridLoss_MWh = transmission + interconnector + distribution losses. Production + NetImports = Load by construction of the settlement. "
             f"Re-reads the last {REVISION_DAYS} days each run (settlement is revised).",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": daily}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
