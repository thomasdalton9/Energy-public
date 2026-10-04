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
  - months Eurostat has published (about 2.5 months behind): the Eurostat monthly total of each fuel (and of load, where taken from
    Eurostat) with the day-to-day shape of the same fuel in the ENTSO-E workbook (daily_shape.reshape: daily = ENTSO-E day x Eurostat
    month / ENTSO-E month; if ENTSO-E has no usable shape for that month, ENTSO-E total generation, else an even spread);
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
import daily_shape  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URL = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem"
# country -> (Eurostat geo, ENTSO-E workbook, output workbook, load from Eurostat too?)
# Germany keeps ENTSO-E's load (its balance closes at 98-99.6% with Eurostat generation). Italy, Poland, Bulgaria and Romania take the
# load from Eurostat's monthly balance (nrg_cb_em) as well: ENTSO-E generation and load both omit embedded/self-consumed power or use
# a gross/net definition that differs from the generation feed, so only a generation + load pair from ONE statistic is consistent.
COUNTRIES = {"Germany": ("DE", "germany_power_generation_daily.xlsx", "germany_eurostat_power_daily.xlsx", False),
             "Italy": ("IT", "italy_power_generation_daily.xlsx", "italy_eurostat_power_daily.xlsx", True),
             "Poland": ("PL", "poland_power_generation_daily.xlsx", "poland_eurostat_power_daily.xlsx", True),
             "Bulgaria": ("BG", "bulgaria_power_generation_daily.xlsx", "bulgaria_eurostat_power_daily.xlsx", True),
             "Romania": ("RO", "romania_power_generation_daily.xlsx", "romania_eurostat_power_daily.xlsx", True)}
BAL_URL = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_em"
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


def fetch_load(geo, since):
    """Monthly consumption incl. network losses (GWh) from nrg_cb_em: 'Available to internal market' (AIM) + distribution losses (DL)."""
    r = requests.get(BAL_URL, params={"format": "JSON", "lang": "EN", "geo": geo, "unit": "GWH", "sinceTimePeriod": since,
                                      "nrg_bal": ["AIM", "DL"], "siec": "E7000"}, headers=H, timeout=120)
    r.raise_for_status()
    s = jsonstat(r.json()).dropna().reset_index(name="v")
    wide = s.pivot_table(index="time", columns="nrg_bal", values="v", aggfunc="sum")
    wide.index = pd.to_datetime(wide.index + "-01")
    wide = wide[wide["AIM"].notna()]
    wide["DL"] = wide["DL"] if "DL" in wide else 0.0
    wide["Load_GWh"] = wide["AIM"] + wide["DL"].fillna(0)
    return wide.sort_index()


def build(geo, since, entsoe_path, eu_load=False, name=""):
    mon = fetch(geo, since)
    # a month counts only when Eurostat has the TOTAL and the main lines (partial months are dropped)
    mon = mon[mon["TOTAL"].notna() & mon["C0000"].notna() & mon["G3000"].notna()]
    ent = pd.read_excel(entsoe_path, sheet_name="Daily")
    ent["date"] = pd.to_datetime(ent["date"])
    ent = ent.set_index("date").sort_index().apply(pd.to_numeric, errors="coerce")
    lmon = None
    if eu_load:
        lmon = fetch_load(geo, since)
        mon = mon[mon.index.isin(lmon.index)]                # a month counts only when generation AND load are both published
    last_month = mon.index.max()
    cut = last_month + pd.offsets.MonthEnd(0)                # last day covered by Eurostat
    days = pd.date_range(mon.index.min(), ent.index.max(), name="date")
    daily = pd.DataFrame(index=days)
    log, devs = [], []
    ent_total = pd.to_numeric(ent[list(MAP)].sum(axis=1, min_count=1), errors="coerce") if all(c in ent for c in MAP) else ent.get("Total_MWh")
    for col, sieces in MAP.items():
        # official Eurostat month total (MWh) with the day-to-day shape of the same fuel in ENTSO-E (fallback: ENTSO-E total generation)
        off = mon.reindex(columns=sieces).fillna(0).sum(axis=1) * 1000.0
        d = daily_shape.reshape(off, ent[col] if col in ent else None, ent_total, label=f"{name} {col}", log=log)
        daily[col] = d.reindex(days)
        devs.append(daily_shape.check_monthly(d, off, col))
    for c in ENTSOE_ONLY:
        daily[c] = ent[c].reindex(days) if c in ent else 0.0
    if eu_load:                                              # load: Eurostat month total with the ENTSO-E load shape, ENTSO-E after the cut
        offl = lmon["Load_GWh"].reindex(mon.index) * 1000.0
        dl = daily_shape.reshape(offl, ent["Load_MWh"], None, label=f"{name} Load", log=log)
        daily["Load_MWh"] = dl.reindex(days).where(days <= cut, ent["Load_MWh"].reindex(days))
        devs.append(daily_shape.check_monthly(dl, offl, "Load"))
    late = days > cut
    for c in MAP:                                            # days after Eurostat's last month: ENTSO-E as published
        daily.loc[late, c] = ent[c].reindex(days[late]) if c in ent else float("nan")
    daily["Total_MWh"] = daily[list(MAP)].sum(axis=1, min_count=1)
    daily = daily.dropna(subset=["Load_MWh"]).fillna({"Gas_MWh": 0})
    daily_shape.print_log(log)
    print(f"  {name}: monthly totals equal Eurostat after reshaping; max deviation {max(devs):.2e} MWh (before rounding)", flush=True)
    assert max(devs) < 1e-3, f"{name}: monthly totals differ from Eurostat by {max(devs)} MWh"
    lastm = daily.loc[last_month:cut, "Total_MWh"]
    nxt = daily.loc[cut + pd.Timedelta(days=1):, "Total_MWh"].head(30)
    if len(lastm) and len(nxt):
        print(f"  {name}: mean daily Total_MWh last scaled month {lastm.mean():.0f} / first 30 unscaled ENTSO-E days {nxt.mean():.0f} = {lastm.mean()/nxt.mean():.2f}", flush=True)
    if lmon is not None:
        mon = mon.join(lmon.rename(columns={"AIM": "AIM_load", "DL": "DL_losses"}), how="left")
    return daily[LAYOUT].round(1), mon, last_month


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01")
    ap.add_argument("--country", default="", help="only this country (default: all)")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    for name, (geo, ent_file, out_file, eu_load) in COUNTRIES.items():
        if args.country and name != args.country:
            continue
        daily, mon, last_month = build(geo, args.start, os.path.join(args.out_dir, ent_file), eu_load, name)
        ann = (daily.groupby(daily.index.year).sum() / 1e6).round(1)
        print(f"{name} TWh per year:\n" + ann.T.to_string(), flush=True)
        cut = (last_month + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
        lines = [f"{name} - monthly net electricity generation by source (Eurostat nrg_cb_pem), spread over the days of each month", "",
                 "Source",
                 "Eurostat 'Net electricity generation by type of fuel - monthly data' (nrg_cb_pem), https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_pem . "
                 "National statistics offices' / TSOs' monthly figures compiled by Eurostat (Germany: Destatis, all producers incl. industrial "
                 "self-generation and rooftop PV; Italy: Terna; Poland, Bulgaria, Romania: their statistics offices), which cover generation that the ENTSO-E per-type feed lacks. Free, no key.",
                 *(["Load: Eurostat nrg_cb_em monthly balance, 'Available to internal market' + distribution losses (consumption incl. network losses), "
                    "https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_em - the same statistic as the generation, so the pair is consistent "
                    "(ENTSO-E's load and generation use different definitions here)."] if eu_load else []),
                 f"ENTSO-E Transparency Platform ({ent_file}) for pumped storage, batteries, load and the days after Eurostat's last month.",
                 "", "Units and definitions",
                 f"MWh per day for dates to {cut} = Eurostat monthly figure (GWh) x the ENTSO-E daily shape of the same fuel (monthly totals equal Eurostat exactly; ENTSO-E total-generation or even spread where ENTSO-E has no usable shape; scale outside 0.2-5 also spreads evenly). Net generation (after the power stations' own use): "
                 "Coal = solid fossil fuels (incl. derived gases), Gas, Oil, Nuclear, Hydro = natural-inflow hydro (pumped-storage generation excluded), Wind, Solar, "
                 "Bioenergy = renewable combustible fuels, Other = non-renewable waste and other combustibles, geothermal, other. "
                 f"From {(last_month + pd.offsets.MonthEnd(0) + pd.Timedelta(days=1)):%Y-%m-%d} (Eurostat is about 2.5 months behind) the fuel columns are the ENTSO-E daily values, "
                 "so those days are on the narrower ENTSO-E definition (supply/load there is the ENTSO-E ratio, not the Eurostat one). PumpedStorage, PumpedStorageConsumption, "
                 "Storage and StorageCharging are ENTSO-E throughout" + ("; Load_MWh is the Eurostat monthly total with the ENTSO-E load shape to the same date, ENTSO-E after it." if eu_load else ", and so is Load_MWh.") + " Eurostat figures are revised; the table is re-read whole each run.",
                 "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; Eurostat {mon.index.min():%Y-%m} to {last_month:%Y-%m}"]
        monthly = mon.copy()
        monthly.index.name = "date"
        xlsx_notes.write_workbook(os.path.join(args.out_dir, out_file), {"Daily": daily, "Monthly (GWh)": monthly}, lines,
                                  {"Source", "Units and definitions", "Last pull"})
        print(f"saved {out_file}")


if __name__ == "__main__":
    main()
