"""
Netherlands electricity production by source from CBS StatLine 84575NED 'Elektriciteitsbalans; aanbod en verbruik' (monthly, mln kWh = GWh),
as a workbook in the same daily layout as the ENTSO-E country workbooks:

  output/Data and Chart Outputs/netherlands_cbs_power_daily.xlsx
    sheet "Daily":   date, <Fuel>_MWh (CBS monthly totals with the ENTSO-E daily shape), Total_MWh, Load_MWh (CBS net consumption + distribution losses)
    sheet "Monthly (GWh)": the CBS columns as published (net production by source, imports/exports by country, losses, net consumption)
    sheet "Units": source and definitions

Why: ENTSO-E's Dutch per-type feed holds under 1 TWh of solar a year (CBS: about 28 TWh in 2025 including rooftop), so the Dutch supply/demand
balance was 15% short. CBS counts all Dutch production (net of the power stations' own use); the table is small, re-read whole each run (revised).

Usage: python3 NETHERLANDS_CBS_POWER.py [--out-dir DIR] [--start 2021-01-01]
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
import daily_shape  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "netherlands_cbs_power_daily.xlsx"
ENTSOE_FILE = "netherlands_power_generation_daily.xlsx"
URL = "https://opendata.cbs.nl/ODataApi/odata/84575NED/TypedDataSet"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept": "application/json"}
# CBS column -> daily workbook column (net production by source)
MAP = {"Kernenergie_4": "Nuclear_MWh", "Kolen_6": "Coal_MWh", "Olieproducten_7": "Oil_MWh", "Aardgas_8": "Gas_MWh",
       "Biomassa_9": "Bioenergy_MWh", "Waterkracht_11": "Hydro_MWh", "WindenergieTotaal_12": "Wind_MWh", "Zonnestroom_15": "Solar_MWh"}
OTHER = ["OverigeBrandstoffenNietHernieuwbaar_10", "OverigeBronnen_16"]
LAYOUT = ["Hydro_MWh", "PumpedStorage_MWh", "Gas_MWh", "Coal_MWh", "Oil_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh", "Bioenergy_MWh",
          "Other_MWh", "Storage_MWh", "Total_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh", "Load_MWh"]


def fetch():
    rows, url = [], URL + "?$format=json&$top=2000"
    while url:
        j = requests.get(url, headers=H, timeout=120).json()
        rows += j.get("value", [])
        url = j.get("odata.nextLink")
    df = pd.DataFrame(rows)
    df = df[df["Perioden"].str.contains("MM")].copy()
    df["month"] = pd.to_datetime(df["Perioden"].str.replace("MM", "-", regex=False) + "-01", format="%Y-%m-%d")
    return df.drop(columns=["ID", "Perioden"]).set_index("month").sort_index().apply(pd.to_numeric, errors="coerce")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    mon = fetch()
    mon = mon[mon.index >= args.start]
    mon = mon[mon["NettoProductieTotaal_3"].notna()]
    days = pd.date_range(mon.index.min(), mon.index.max() + pd.offsets.MonthEnd(0), name="date")
    daily = pd.DataFrame(index=days)
    # Each CBS month keeps its official total exactly; the day-to-day shape comes from the ENTSO-E Dutch workbook (same fuel; fallback
    # ENTSO-E total generation, then an even spread) - see daily_shape.py. Scale factors outside 0.2-100 spread evenly (CBS solar is 20-40x
    # ENTSO-E's under-reporting, so the solar bound is wide: the ENTSO-E shape is still the right solar profile).
    ent = pd.read_excel(os.path.join(args.out_dir, ENTSOE_FILE), sheet_name="Daily")
    ent["date"] = pd.to_datetime(ent["date"])
    ent = ent.set_index("date").sort_index().apply(pd.to_numeric, errors="coerce")
    ent_total = ent["Total_MWh"] if "Total_MWh" in ent else None
    log, devs, offs = [], [], {}
    series = {dst: mon[src].fillna(0) * 1000.0 for src, dst in MAP.items()}
    series["Other_MWh"] = mon[OTHER].fillna(0).sum(axis=1) * 1000.0
    series["Load_MWh"] = (mon["NettoVerbruikBerekend_30"] + mon["Distributieverliezen_29"]) * 1000.0
    for dst, off in series.items():
        is_load = dst == "Load_MWh"
        d = daily_shape.reshape(off, ent[dst] if dst in ent else None, None if is_load else ent_total, hi=100.0 if dst == "Solar_MWh" else 5.0,
                                label=f"Netherlands {dst}", log=log)
        daily[dst] = d.reindex(days)
        devs.append(daily_shape.check_monthly(d, off, dst))
        offs[dst] = off
    daily_shape.print_log(log)
    print(f"monthly totals equal CBS after reshaping; max deviation {max(devs):.2e} MWh (before rounding)", flush=True)
    assert max(devs) < 1e-3, f"monthly totals differ from CBS by {max(devs)} MWh"
    daily["Total_MWh"] = daily[[c for c in MAP.values()] + ["Other_MWh"]].sum(axis=1)
    for c in ("PumpedStorage_MWh", "Storage_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh"):
        daily[c] = 0.0
    # Load = CBS calculated net consumption + distribution losses (grid demand). ENTSO-E's Dutch load agrees within 1% in 2024-25 but is about
    # 10% below it in 2021-22 (traced: net imports and production match CBS, only the load differs), so CBS is used for the whole series.
    load_note = ("Load_MWh = CBS calculated net consumption plus distribution losses (production + imports - exports, so the balance against "
                 "production and flows closes by construction; ENTSO-E's Dutch load agrees within 1% in 2024-25 but is about 10% lower in 2021-22)")
    daily = daily[LAYOUT].round(1)
    ann = (daily.groupby(daily.index.year).sum() / 1e6).round(1)
    print("TWh per year:\n" + ann.T.to_string(), flush=True)
    lines = ["Netherlands - monthly electricity production by source (CBS StatLine 84575NED), with the ENTSO-E daily shape", "",
             "Source", "Statistics Netherlands (CBS) 'Elektriciteitsbalans; aanbod en verbruik', https://opendata.cbs.nl/ODataApi/odata/84575NED . Free, no key.",
             "", "Units and definitions",
             "MWh per day = CBS monthly figure (mln kWh) with the ENTSO-E daily shape of the same fuel (monthly totals equal CBS exactly; ENTSO-E total generation or an even spread where ENTSO-E has no usable shape). Net production by source (after the power stations' own use): "
             "Nuclear, Coal, Oil, Gas, Bioenergy (biomass), Hydro, Wind (land + sea), Solar (all PV including rooftop), Other (non-renewable "
             "other fuels + other sources). " + load_note + ". The 'Monthly (GWh)' sheet keeps the CBS columns as published, including imports/exports "
             "by country and distribution losses. CBS figures are revised; the table is re-read whole each run.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {mon.index.min():%Y-%m} to {mon.index.max():%Y-%m}"]
    monthly = mon.copy()
    monthly.index.name = "date"
    xlsx_notes.write_workbook(os.path.join(args.out_dir, FILE), {"Daily": daily, "Monthly (GWh)": monthly}, lines,
                              {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
