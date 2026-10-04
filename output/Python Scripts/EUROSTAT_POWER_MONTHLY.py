"""
Germany electricity production by source from Eurostat nrg_cb_pem 'Net electricity generation by type of fuel - monthly data' (GWh;
for Germany the figures are Destatis' monthly electricity statistics, which cover ALL producers - public utilities, industrial
self-generators and rooftop PV), as a workbook in the same daily layout as the ENTSO-E country workbooks:

  output/Data and Chart Outputs/germany_eurostat_power_daily.xlsx
    sheet "Daily":   date, <Fuel>_MWh, Total_MWh, Load_MWh
    sheet "Monthly (GWh)": Eurostat columns as published (siec codes)
    sheet "Units": source and definitions

Why: ENTSO-E's German per-type feed holds only the TSO-visible plants. Industrial self-generation and part of the small PV and
biomass fleet are missing (Eurostat 2024: 441 TWh net generation vs 429 TWh in the ENTSO-E feed), so the German supply/demand balance
closed at 95-98%.

How the daily layout is filled:
  - months Eurostat has published (about 2.5 months behind): each Eurostat month spread evenly over its days, by fuel;
  - later days: the ENTSO-E daily values, unchanged, so the series stays current (the cut-over is on the Units sheet);
  - pumped-storage output and consumption, battery columns and the load always come from the ENTSO-E workbook (Eurostat
    reports no pumping consumption; its hydro lines exclude pumped-storage generation to avoid double counting).

Usage: python3 EUROSTAT_POWER_MONTHLY.py [--out-dir DIR] [--start 2021-01]
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
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URL = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem"
# country -> (Eurostat geo, ENTSO-E workbook, output workbook)
COUNTRIES = {"Germany": ("DE", "germany_power_generation_daily.xlsx", "germany_eurostat_power_daily.xlsx")}
# Eurostat siec -> workbook column (sums). RA130 = pumped hydro (left out: pumped storage comes from ENTSO-E), RA110 + RA120 = natural-inflow hydro.
MAP = {"Hydro_MWh": ["RA110", "RA120"], "Gas_MWh": ["G3000"], "Coal_MWh": ["C0000"], "Oil_MWh": ["O4000XBIO"], "Nuclear_MWh": ["N9000"],
       "Wind_MWh": ["RA300"], "Solar_MWh": ["RA400"], "Bioenergy_MWh": ["CF_R"],
       "Other_MWh": ["CF_NR_OTH", "X9900", "RA200", "RA500_5160"]}
ENTSOE_ONLY = ["PumpedStorage_MWh", "Storage_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh", "Load_MWh"]
LAYOUT = ["Hydro_MWh", "PumpedStorage_MWh", "Gas_MWh", "Coal_MWh", "Oil_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh", "Bioenergy_MWh",
          "Other_MWh", "Storage_MWh", "Total_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh", "Load_MWh"]
NEED = sorted({s for v in MAP.values() for s in v})


def jsonstat(j):
    ids = j["id"]
    cats = []
    for d in ids:
        idx = j["dimension"][d]["category"]["index"]
        cats.append(list(idx.keys()) if isinstance(idx, dict) else list(idx))
    s = pd.Series(float("nan"), index=pd.MultiIndex.from_product(cats, names=ids), dtype=float)
    vals = j["value"]
    items = vals.items() if isinstance(vals, dict) else enumerate(vals)
    for k, v in items:
        if v is not None:
            s.iloc[int(k)] = v
    return s


def fetch(geo, since):
    r = requests.get(URL, params={"format": "JSON", "lang": "EN", "geo": geo, "unit": "GWH", "sinceTimePeriod": since}, headers=H, timeout=120)
    r.raise_for_status()
    s = jsonstat(r.json()).dropna().reset_index(name="v")
    wide = s.pivot_table(index="time", columns="siec", values="v", aggfunc="sum")
    wide.index = pd.to_datetime(wide.index + "-01")
    return wide.sort_index()


def build(geo, since, entsoe_path):
    mon = fetch(geo, since)
    # a month counts only when Eurostat has the TOTAL and the main lines (partial months are dropped)
    mon = mon[mon["TOTAL"].notna() & mon["C0000"].notna() & mon["G3000"].notna()]
    ent = pd.read_excel(entsoe_path, sheet_name="Daily")
    ent["date"] = pd.to_datetime(ent["date"])
    ent = ent.set_index("date").sort_index().apply(pd.to_numeric, errors="coerce")
    last_month = mon.index.max()
    cut = last_month + pd.offsets.MonthEnd(0)                # last day covered by Eurostat
    days = pd.date_range(mon.index.min(), ent.index.max(), name="date")
    daily = pd.DataFrame(index=days)
    key = days.to_period("M").to_timestamp()
    dim = pd.Series(days.days_in_month, index=days)
    for col, sieces in MAP.items():
        gwh = mon.reindex(columns=sieces).fillna(0).sum(axis=1)
        daily[col] = gwh.reindex(key).to_numpy() * 1000.0 / dim.to_numpy()
    for c in ENTSOE_ONLY:
        daily[c] = ent[c].reindex(days) if c in ent else 0.0
    late = days > cut
    for c in MAP:                                            # days after Eurostat's last month: ENTSO-E as published
        daily.loc[late, c] = ent[c].reindex(days[late]) if c in ent else float("nan")
    daily["Total_MWh"] = daily[list(MAP)].sum(axis=1, min_count=1)
    daily = daily.dropna(subset=["Load_MWh"]).fillna({"Gas_MWh": 0})
    return daily[LAYOUT].round(1), mon, last_month


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    for name, (geo, ent_file, out_file) in COUNTRIES.items():
        daily, mon, last_month = build(geo, args.start, os.path.join(args.out_dir, ent_file))
        ann = (daily.groupby(daily.index.year).sum() / 1e6).round(1)
        print(f"{name} TWh per year:\n" + ann.T.to_string(), flush=True)
        cut = (last_month + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
        lines = [f"{name} - monthly net electricity generation by source (Eurostat nrg_cb_pem), spread over the days of each month", "",
                 "Source",
                 "Eurostat 'Net electricity generation by type of fuel - monthly data' (nrg_cb_pem), https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_pem . "
                 "For Germany the figures come from Destatis' monthly electricity statistics and cover all producers, including industrial "
                 "self-generation and rooftop PV that the ENTSO-E per-type feed lacks. Free, no key.",
                 "ENTSO-E Transparency Platform (germany_power_generation_daily.xlsx) for pumped storage, batteries, load and the days after Eurostat's last month.",
                 "", "Units and definitions",
                 f"MWh per day = Eurostat monthly figure (GWh) / days in the month, for dates to {cut}. Net generation (after the power stations' own use): "
                 "Coal = solid fossil fuels (incl. derived gases), Gas, Oil, Nuclear, Hydro = natural-inflow hydro (pumped-storage generation excluded), Wind, Solar, "
                 "Bioenergy = renewable combustible fuels, Other = non-renewable waste and other combustibles, geothermal, other. "
                 f"From {(last_month + pd.offsets.MonthEnd(0) + pd.Timedelta(days=1)):%Y-%m-%d} (Eurostat is about 2.5 months behind) the fuel columns are the ENTSO-E daily values, "
                 "so those days are on the narrower ENTSO-E definition (about 3% lower for Germany). PumpedStorage, PumpedStorageConsumption, "
                 "Storage, StorageCharging and Load_MWh are ENTSO-E throughout. Eurostat figures are revised; the table is re-read whole each run.",
                 "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; Eurostat {mon.index.min():%Y-%m} to {last_month:%Y-%m}"]
        monthly = mon.copy()
        monthly.index.name = "date"
        xlsx_notes.write_workbook(os.path.join(args.out_dir, out_file), {"Daily": daily, "Monthly (GWh)": monthly}, lines,
                                  {"Source", "Units and definitions", "Last pull"})
        print(f"saved {out_file}")


if __name__ == "__main__":
    main()
