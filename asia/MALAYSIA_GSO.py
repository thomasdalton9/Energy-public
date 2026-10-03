"""
Malaysia (Peninsular grid) electricity from GSO, the Grid System Operator (Single Buyer / Energy
Commission), https://www.gso.org.my/SystemData/CurrentGen.aspx. Found via
discovery_archive/asia/SEA_DISCOVERY_MY_ID_BN.py.

The GSO pages load their charts from ASP.NET page methods that take a date range (POST JSON
{"Fromdate": "dd/mm/yyyy", "Todate": "dd/mm/yyyy"}, no key):
  CurrentGen.aspx/GetChartDataSource    10-minute generation by fuel, MW: Coal, Gas, CoGen, Oil, Hydro,
                                        Solar (large-scale solar on the grid). History back to 2020.
  SystemDemand.aspx/GetChartDataSource  10-minute system demand, MW
  PowerStation.aspx/GetDataSource       plant list: name, fuel, type, capacity MW, PPA expiry (current only)

Writes two workbooks:
  malaysia_power_generation_daily.xlsx  'Daily' (standard layout, MWh per day = mean of the day's 10-minute MW
                                        x 24, so a day missing a few intervals is not understated), 'Demand'
                                        (daily average and peak MW)
  malaysia_power_prices.xlsx            'Daily' system marginal price (SMP) from Single Buyer: daily average, max
                                        and min of the half-hourly SMP, RM/MWh; 'Half-hourly' (last 120 days).
                                        Single Buyer's API returns the last 12 months on every call, so each run
                                        adds to the saved history.
  malaysia_power_capacity.xlsx          'Monthly' (standard layout: installed MW by fuel, one snapshot per month
                                        from the plant list - the list holds only today's fleet, so history
                                        starts with the first run), 'Plants' (latest list)

Peninsular Malaysia only (about 80% of national demand); Sabah and Sarawak run separate grids with no
public daily data. CoGen (gas-fired cogeneration) is counted as Gas.

Incremental: the Daily sheet is the history store; only days not saved yet (from 2021) are fetched, plus
the last REVISION_DAYS days again. Runs on the 1st and 15th.

    python3 asia/MALAYSIA_GSO.py
"""
import argparse
import json
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import xlsx_notes  # noqa: E402
import power_capacity_std as cap_std  # noqa: E402

GSO = "https://www.gso.org.my/SystemData"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Content-Type": "application/json; charset=utf-8", "Origin": "https://www.gso.org.my",
     "X-Requested-With": "XMLHttpRequest"}
T = (15, 120)
DATA_START = date(2021, 1, 1)
CHUNK_DAYS = 7
REVISION_DAYS = 18   # runs are 14-17 days apart: re-read everything since the last run, with 1-2 days spare (provisional days get final)
OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
GEN_OUT = os.path.join(OUT_DIR, "malaysia_power_generation_daily.xlsx")
CAP_OUT = os.path.join(OUT_DIR, "malaysia_power_capacity.xlsx")
PRICE_OUT = os.path.join(OUT_DIR, "malaysia_power_prices.xlsx")
SMP_API = "https://www.singlebuyer.com.my/api/v1/smp/actual-forecast"
# GSO series -> standard fuel
FUEL_MAP = {"Coal": "Coal", "Gas": "Gas", "CoGen": "Gas", "Oil": "Oil", "Hydro": "Hydro", "Solar": "Solar"}
FUELS = ["Hydro", "Gas", "Solar", "Coal", "Oil", "Other"]
# plant list 'Fuel' -> standard capacity fuel
PLANT_FUEL = {"water": "Hydro", "hydro": "Hydro", "gas": "Gas", "coal": "Coal", "solar": "Solar", "oil": "Oil",
              "distillate": "Oil", "diesel": "Oil", "biomass": "Bioenergy", "biogas": "Bioenergy", "cogen": "Gas"}


def out(*a):
    print(*a, flush=True)


def post(page, method, d0, d1):
    body = json.dumps({"Fromdate": d0.strftime("%d/%m/%Y"), "Todate": d1.strftime("%d/%m/%Y")})
    for i in range(4):
        try:
            r = requests.post(f"{GSO}/{page}/{method}", data=body, headers=dict(H, Referer=f"{GSO}/{page}"),
                              timeout=T)
            r.raise_for_status()
            d = r.json().get("d", [])
            return json.loads(d) if isinstance(d, str) else d
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                raise
            out(f"  retry {page} {d0}: {e}")
            time.sleep(5 * (i + 1))


def daily_from(records, cols):
    """10-minute records -> per day: mean MW x 24 (MWh), plus the number of intervals seen."""
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df["DT"] = pd.to_datetime(df["DT"])
    df = df.set_index("DT")[[c for c in cols if c in df]].apply(pd.to_numeric, errors="coerce")
    g = df.groupby(df.index.normalize())
    return g.mean().mul(24), g.size()


def fetch_days(days):
    """Generation by fuel (MWh/day) and demand (avg/peak MW) for the given dates, CHUNK_DAYS per request."""
    gen, dem = [], []
    days = sorted(days)
    i = 0
    while i < len(days):
        d0 = days[i]
        d1 = min(d0 + timedelta(days=CHUNK_DAYS - 1), days[-1])
        g = post("CurrentGen.aspx", "GetChartDataSource", d0, d1)
        if g:
            mwh, n = daily_from(g, list(FUEL_MAP))
            std = pd.DataFrame(index=mwh.index)
            for src, fuel in FUEL_MAP.items():
                if src in mwh:
                    std[f"{fuel}_MWh"] = std.get(f"{fuel}_MWh", 0) + mwh[src].fillna(0)
            std["Intervals"] = n
            gen.append(std)
        s = post("SystemDemand.aspx", "GetChartDataSource", d0, d1)
        if s:
            df = pd.DataFrame(s)
            df["DT"] = pd.to_datetime(df["DT"])
            mw = pd.to_numeric(df.set_index("DT")["MW"], errors="coerce")
            gd = mw.groupby(mw.index.normalize())
            dem.append(pd.DataFrame({"Demand_avg_MW": gd.mean().round(0), "Demand_peak_MW": gd.max(),
                                     "Demand_min_MW": gd.min()}))
        out(f"  {d0}..{d1}: {sum(len(x) for x in gen)} generation days so far")
        while i < len(days) and days[i] <= d1:
            i += 1
        time.sleep(0.5)
    gen = pd.concat(gen) if gen else pd.DataFrame()
    dem = pd.concat(dem) if dem else pd.DataFrame()
    want = set(pd.Timestamp(d) for d in days)
    gen = gen[gen.index.isin(want)] if not gen.empty else gen
    dem = dem[dem.index.isin(want)] if not dem.empty else dem
    return gen, dem


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    if old.empty:
        return new.sort_index()
    if new.empty:
        return old
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def generation(gen_out, start):
    old_gen, old_dem = read_sheet(gen_out, "Daily"), read_sheet(gen_out, "Demand")
    yesterday = date.today() - timedelta(days=1)
    have = set(old_gen.index.date) if not old_gen.empty else set()
    revise = {yesterday - timedelta(days=k) for k in range(REVISION_DAYS)}
    todo = [d for d in (start + timedelta(days=k) for k in range((yesterday - start).days + 1))
            if d not in have or d in revise]
    out(f"Generation: {len(have)} days saved; fetching {len(todo)} "
        f"({todo[0] if todo else '-'}..{todo[-1] if todo else '-'})")
    gen, dem = fetch_days(todo) if todo else (pd.DataFrame(), pd.DataFrame())
    if not gen.empty:
        for f in FUELS:
            gen[f"{f}_MWh"] = gen.get(f"{f}_MWh", 0.0)
        gen = gen[[f"{f}_MWh" for f in FUELS] + ["Intervals"]]
        gen.insert(len(FUELS), "Total_MWh", gen[[f"{f}_MWh" for f in FUELS]].sum(axis=1))
        gen = gen.round(1)
    daily = merge(old_gen, gen)
    demand = merge(old_dem, dem)
    daily.index.name = demand.index.name = "date"
    if daily.empty:
        raise SystemExit("No GSO generation data")
    notes = [
        "UNITS",
        "Daily: MWh per day by fuel = mean of the day's 10-minute generation (MW) x 24. Total_MWh = sum of fuels. "
        "Intervals = number of 10-minute values GSO returned for the day (144 for a full day; fewer means "
        "missing intervals, which the mean x 24 bridges).",
        "Demand: Peninsular system demand, MW: daily average, peak and minimum of the 10-minute values.",
        "",
        "COVERAGE",
        f"Peninsular Malaysia grid only (Sabah and Sarawak are separate systems). Daily from {daily.index.min():%Y-%m-%d} "
        f"to {daily.index.max():%Y-%m-%d}.",
        "Fuels as GSO reports them: Coal, Gas, CoGen (gas cogeneration, counted as Gas here), Oil, Hydro, Solar "
        "(large-scale solar connected to the grid; rooftop solar is not metered by GSO). Other_MWh is kept for the "
        "standard layout and is zero.",
        "",
        "SOURCE",
        "GSO (Grid System Operator, Malaysia), System Data - Generation Mix and System Demand: "
        "https://www.gso.org.my/SystemData/CurrentGen.aspx (page methods CurrentGen.aspx/GetChartDataSource and "
        "SystemDemand.aspx/GetChartDataSource, POST with a date range).",
    ]
    xlsx_notes.write_workbook(gen_out, {"Daily": daily, "Demand": demand}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {gen_out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


def capacity(cap_out):
    t = date.today()
    plants = pd.DataFrame(post("PowerStation.aspx", "GetDataSource", t - timedelta(days=1), t))
    if plants.empty:
        out("Plant list empty - capacity not updated")
        return
    plants["Fuel"] = plants["Fuel"].astype(str).str.strip()
    plants["Capacity_MW"] = pd.to_numeric(plants.get("Capacity (MW)", plants.get("Capacity")), errors="coerce")
    plants["Standard_fuel"] = [next((v for k, v in PLANT_FUEL.items() if k in f"{f} {t}".lower()), "Other")
                               for f, t in zip(plants["Fuel"], plants.get("Type", ""))]   # GSO lists hydro as 'Water'

    by = plants.groupby("Standard_fuel")["Capacity_MW"].sum()
    month = pd.Timestamp(t.replace(day=1))
    row = cap_std.standard(pd.DataFrame([by.to_dict()], index=[month]))
    monthly = cap_std.load_monthly(cap_out)
    monthly = merge(monthly, row)
    monthly.index.name = "date"
    keep = [c for c in ("Name", "Fuel", "Type", "Capacity_MW", "Standard_fuel", "PPA/SLA Expiry Year") if c in plants]
    notes = [
        "UNITS",
        "Monthly: installed capacity, MW, by fuel, Peninsular Malaysia, as listed in GSO's power station list on the "
        "run date (one row per month: the latest run in that month). Total_MW = sum of fuels.",
        "Plants: the latest list (licensed generating plants dispatched by GSO, with PPA/SLA expiry).",
        "",
        "COVERAGE",
        f"GSO publishes only the current fleet, so the monthly history starts with this pull's first run "
        f"({monthly.index.min():%b %Y}). Peninsular Malaysia only; large-scale solar included, rooftop solar "
        "and plants not dispatched by GSO are not.",
        "",
        "SOURCE",
        "GSO (Grid System Operator, Malaysia), Power Stations: https://www.gso.org.my/SystemData/PowerStation.aspx",
    ]
    cap_std.write(cap_out, monthly, {"Plants": plants[keep]}, notes, {"UNITS", "COVERAGE", "SOURCE"})


def prices(price_out):
    """Single Buyer half-hourly SMP (RM/kWh in the API) -> RM/MWh; the API answers with the last 12 months of
    actuals (its date parameters are ignored), merged into the saved history."""
    r = requests.get(SMP_API, params={"from": date.today().isoformat()},
                     headers={"User-Agent": H["User-Agent"], "Accept": "application/json",
                              "Referer": "https://www.singlebuyer.com.my/"}, timeout=T)
    r.raise_for_status()
    j = r.json()
    act = (j.get("meta") or {}).get("data", {}).get("actual") or (j.get("data") or {}).get("actual") or []
    hh = pd.DataFrame(act)
    if hh.empty:
        out("No SMP actuals returned")
        return
    hh["t"] = pd.to_datetime(hh["t"])
    hh = hh.set_index("t")["v"].astype(float).mul(1000).rename("SMP_RM_per_MWh").to_frame()
    old_hh = read_sheet(price_out, "Half-hourly")
    hh = merge(old_hh, hh)
    g = hh["SMP_RM_per_MWh"].groupby(hh.index.normalize())
    new = pd.DataFrame({"SMP_avg_RM_per_MWh": g.mean(), "SMP_max_RM_per_MWh": g.max(), "SMP_min_RM_per_MWh": g.min(),
                        "Periods": g.size()}).round(2)
    daily = merge(read_sheet(price_out, "Daily"), new[new["Periods"] >= 46])   # complete days only
    hh = hh[hh.index >= hh.index.max() - pd.Timedelta(days=120)]
    daily.index.name, hh.index.name = "date", "time"
    notes = [
        "UNITS",
        "RM/MWh (Malaysian ringgit per MWh; the API gives RM/kWh, x1,000). Daily: average, max and min of the 48 "
        "half-hourly system marginal prices; Periods = half-hours present. Half-hourly: last 120 days.",
        "",
        "COVERAGE",
        f"Peninsular Malaysia. Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. Single Buyer's "
        "API holds a rolling 12 months; earlier days are kept from previous runs.",
        "",
        "SOURCE",
        "Single Buyer (Malaysia), System Marginal Price: https://www.singlebuyer.com.my/ "
        "(API https://www.singlebuyer.com.my/api/v1/smp/actual-forecast).",
    ]
    xlsx_notes.write_workbook(price_out, {"Daily": daily, "Half-hourly": hh}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {price_out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gen-out", default=GEN_OUT)
    ap.add_argument("--cap-out", default=CAP_OUT)
    ap.add_argument("--price-out", default=PRICE_OUT)
    ap.add_argument("--start", default=DATA_START.isoformat())
    args = ap.parse_args()
    os.makedirs(os.path.dirname(os.path.abspath(args.gen_out)), exist_ok=True)
    steps = (("Generation", lambda p: generation(p, date.fromisoformat(args.start)), args.gen_out),
             ("Capacity", capacity, args.cap_out), ("Prices", prices, args.price_out))
    for step, fn, path in steps:   # one source down (GSO or Single Buyer) does not stop the others
        try:
            fn(path)
        except Exception as e:  # noqa: BLE001
            out(f"{step} step failed: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
