"""
US installed generating capacity by fuel, monthly, from EIA's generator
inventory (Form EIA-860M) via the EIA API v2:

  https://api.eia.gov/v2/electricity/operating-generator-capacity/data/  (needs EIA_API_KEY)

The API returns one row per generator per month (~28,000 a month); this
sums net summer capacity by energy source for every generator in EIA's
operating inventory (status OP operating, SB standby, OA/OS out of
service - the same set as the EIA-860M "Operating" sheet), all 50 states.

Writes us_power_capacity.xlsx:
  Monthly        standard capacity layout (date, <Fuel>_MW, Total_MW) plus
                 Battery_storage_MW and Pumped_storage_MW (storage is kept
                 out of the fuel columns; Total_MW includes it)
  By technology  MW by EIA technology (e.g. Natural Gas Fired Combined Cycle)

Regional sheets (US_Total, ERCOT, PJM, MISO, SPP, CAISO, NYISO, ISONE, Southern, TVA; the EIA-930 regions used by
NORTH_AMERICA_MASTER): the API has a balancing_authority_code field/facet, so each generator row is assigned to the
EIA-930 balancing authority it is registered to (REGION_BA below) and net summer MW are summed by fuel per region,
same layout as Monthly. US_Total is the Lower 48 (AK and HI left out, as EIA-930 has none). Generators with no BA
code (mainly AK/HI and a few small plants) fall in no region. A "Gas capacity factor" sheet divides EIA-930 gas
generation (eia930_fuel_mix_daily.xlsx, complete months only) by gas capacity x hours.

Incremental: fetches only months not saved yet plus the latest two saved
months (EIA revises the most recent inventory); a new file backfills from
2021. Saves after every month so a timeout keeps what was fetched.

Usage: python3 US_POWER_CAPACITY_EIA.py [--out "output/Data and Chart Outputs/us_power_capacity.xlsx"]
"""
import argparse
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://api.eia.gov/v2/electricity/operating-generator-capacity/data/"
HISTORY_START = "2021-01"
REFRESH_MONTHS = 2
PAGE = 5000
STATUSES = ["OP", "SB", "OA", "OS"]
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "us_power_capacity.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]

# EIA energy source code -> standard fuel (anything not listed -> Other)
SOURCE = {
    "NG": "Gas",
    **{c: "Coal" for c in ("BIT", "SUB", "LIG", "ANT", "RC", "WC", "SGC")},
    **{c: "Oil" for c in ("DFO", "RFO", "JF", "KER", "PC", "WO", "SGP", "PG")},
    "NUC": "Nuclear", "WAT": "Hydro", "SUN": "Solar", "WND": "Wind",
    **{c: "Bioenergy" for c in ("WDS", "WDL", "BLQ", "AB", "MSW", "OBS", "OBL", "OBG", "LFG", "SLW")},
    "MWH": "Battery_storage",
}


# EIA-930 region sheet -> balancing authority code(s) in the EIA-860M inventory
REGION_BA = {"ERCOT": ["ERCO"], "PJM": ["PJM"], "MISO": ["MISO"], "SPP": ["SWPP"], "CAISO": ["CISO"],
             "NYISO": ["NYIS"], "ISONE": ["ISNE"], "Southern": ["SOCO"], "TVA": ["TVA"]}
REGIONS = ["US_Total"] + list(REGION_BA)
FUEL_COLS = [f"{f}_MW" for f in FUELS + ["Battery_storage", "Pumped_storage"]]
EIA930_XLSX = os.path.join(ROOT, "output", "Data and Chart Outputs", "eia930_fuel_mix_daily.xlsx")


def get_page(key, month, offset):
    params = {"api_key": key, "frequency": "monthly", "data[0]": "net-summer-capacity-mw", "start": month,
              "end": month, "offset": offset, "length": PAGE,
              # a stable order, so pages neither overlap nor skip rows
              "sort[0][column]": "plantid", "sort[0][direction]": "asc",
              "sort[1][column]": "generatorid", "sort[1][direction]": "asc"}
    for i, st in enumerate(STATUSES):
        params[f"facets[status][{i}]"] = st
    for attempt in range(4):
        try:
            r = requests.get(URL, params=params, timeout=(10, 120))
            r.raise_for_status()
            return r.json()["response"]
        except (requests.RequestException, ValueError, KeyError) as e:
            print(f"    {month} offset {offset}: attempt {attempt + 1}/4 failed: {type(e).__name__}: {str(e)[:150]}",
                  flush=True)
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(f"EIA API failed for {month} offset {offset}")


def fetch_month(key, month):
    """All generator rows for one month -> (fuel MW Series, technology MW Series)."""
    first = get_page(key, month, 0)
    total = int(first.get("total", 0))
    rows = list(first.get("data", []))
    with ThreadPoolExecutor(max_workers=4) as ex:
        for resp in ex.map(lambda off: get_page(key, month, off), range(PAGE, total, PAGE)):
            rows += resp.get("data", [])
    if not rows:
        return None
    d = pd.DataFrame(rows)
    d["mw"] = pd.to_numeric(d["net-summer-capacity-mw"], errors="coerce").fillna(0)
    d["fuel"] = d["energy_source_code"].map(SOURCE).fillna("Other")
    d.loc[d["technology"].eq("Hydroelectric Pumped Storage"), "fuel"] = "Pumped_storage"
    d.loc[d["technology"].isin(["Batteries", "Flywheels"]), "fuel"] = "Battery_storage"
    if len(d) != total:
        print(f"  WARNING {month}: got {len(d)} rows, API total {total}", flush=True)
    print(f"  {month}: {len(d)} of {total} generator rows, {d['mw'].sum() / 1000:.1f} GW", flush=True)
    ba = d["balancing_authority_code"] if "balancing_authority_code" in d else pd.Series(None, index=d.index)
    regional = {"US_Total": d[~d["stateid"].isin(["AK", "HI"])].groupby("fuel")["mw"].sum()}
    for reg, codes in REGION_BA.items():
        regional[reg] = d[ba.isin(codes)].groupby("fuel")["mw"].sum()
    return (d.groupby("fuel")["mw"].sum(), d.groupby(d["technology"].fillna("Unknown"))["mw"].sum(), regional)


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def latest_month(key):
    r = requests.get(URL.replace("/data/", "/"), params={"api_key": key}, timeout=(10, 60))
    r.raise_for_status()
    return pd.Timestamp(r.json()["response"]["endPeriod"])


def region_row(fuel, m):
    row = pd.DataFrame([fuel]).rename(columns=lambda c: f"{c}_MW")
    row.index = [m]
    row["Total_MW"] = fuel.sum()
    return row


def gas_capacity_factor(regional, path=EIA930_XLSX):
    """Gas CF %, per region: EIA-930 Natural_Gas_MWh in complete months / (gas MW x hours in month)."""
    out = {}
    for reg in REGIONS:
        cap = regional.get(reg)
        if cap is None or cap.empty or "Gas_MW" not in cap:
            continue
        try:
            d = pd.read_excel(path, sheet_name=reg, usecols=["date", "Natural_Gas_MWh"])
        except (FileNotFoundError, ValueError, KeyError, OSError):
            continue
        d["date"] = pd.to_datetime(d["date"].astype(str), errors="coerce")
        d = d.dropna(subset=["date"]).set_index("date")["Natural_Gas_MWh"].astype(float)
        g = d.resample("MS").agg(["sum", "count"])
        g = g[g["count"] >= g.index.days_in_month]          # full months only
        gen = g["sum"]
        mw = cap["Gas_MW"].reindex(gen.index)
        cf = 100.0 * gen / (mw * 24 * gen.index.days_in_month)
        out[f"{reg}_gas_CF_pct"] = cf
        out[f"{reg}_gas_GWh"] = gen / 1000.0
    return pd.DataFrame(out).dropna(how="all")


def save(path, monthly, tech, regional=None, with_cf=True):
    notes = [
        "UNITS",
        "Net summer capacity, MW, at the end of each month (EIA-860M generator inventory).",
        "Monthly: Hydro (conventional hydroelectric), Gas, Wind, Solar (PV and thermal), Coal (incl. waste coal and "
        "coal-derived gas), Nuclear, Oil (distillate, residual, jet fuel, kerosene, petroleum coke, waste oil, "
        "propane), Bioenergy (wood, black liquor, MSW, landfill gas, other biomass), Other (geothermal, other gases, "
        "waste heat, ...). Battery_storage_MW (batteries, flywheels) and Pumped_storage_MW are kept separate; "
        "Total_MW includes them.",
        "By technology: MW by EIA's technology label.",
        "Regional sheets (US_Total = Lower 48; ERCOT, PJM, MISO, SPP, CAISO, NYISO, ISONE, Southern, TVA): the same "
        "columns summed by the balancing authority EIA-860M assigns to each generator (ERCO, PJM, MISO, SWPP, CISO, "
        "NYIS, ISNE, SOCO, TVA). A BA's registered footprint is not exactly its EIA-930 generation area for every plant "
        "but is EIA's own mapping. Generators without a BA code (mainly Alaska and Hawaii) are in no region.",
        "Gas capacity factor: EIA-930 Natural_Gas_MWh (months with every day present) / (Gas_MW x hours in month), %. "
        "Gas_MW is net summer capacity, so summer factors read slightly high; <Region>_gas_GWh is the generation used.",
        "Generators counted: operating inventory - status OP (operating), SB (standby), OA and OS (out of service), "
        "as on the EIA-860M 'Operating' sheet. All 50 states (EIA-860M covers plants of 1 MW and above).",
        "",
        "COVERAGE",
        f"United States, monthly, {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y}. EIA publishes each "
        "month's inventory about two months later.",
        "",
        "SOURCE",
        f"EIA API v2, electricity/operating-generator-capacity (Form EIA-860M): {URL}",
        "https://www.eia.gov/electricity/data/eia860m/",
    ]
    m, t = monthly.copy(), tech.copy()
    for x in (m, t):
        x.index.name = "date"
    sheets = {"Monthly": m.round(1), "By technology": t.round(1)}
    for reg, df in (regional or {}).items():
        if not df.empty:
            x = df.copy()
            x.index.name = "date"
            sheets[reg] = x.round(1)
    if regional and with_cf:
        cf = gas_capacity_factor(regional)
        if not cf.empty:
            cf.index.name = "date"
            sheets["Gas capacity factor"] = cf.round(2)
    xlsx_notes.write_workbook(path, sheets, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-months", type=int, default=100, help="cap on months fetched this run (backfill chunks)")
    args = ap.parse_args()
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise SystemExit("EIA_API_KEY not set")

    monthly, tech = load(args.out, "Monthly"), load(args.out, "By technology")
    regional = {r: load(args.out, r) for r in REGIONS}
    end = latest_month(key)
    months = pd.date_range(HISTORY_START, end, freq="MS")
    have = set(monthly.index) if not monthly.empty else set()
    for r in REGIONS:   # a month is saved only when every regional sheet has it (older files had none: backfill)
        have &= set(regional[r].index)
    todo = [m for m in months if m not in have]
    if have:
        todo += [m for m in sorted(have)[-REFRESH_MONTHS:] if m not in todo]
    todo = sorted(todo)[:args.max_months]
    print(f"EIA inventory runs to {end:%Y-%m}; {len(have)} months saved; fetching {len(todo)}", flush=True)

    done = 0
    for m in todo:
        res = fetch_month(key, m.strftime("%Y-%m"))
        if res is None:
            print(f"  {m:%Y-%m}: no rows", flush=True)
            continue
        fuel, by_tech, by_region = res
        row = pd.DataFrame([fuel]).rename(columns=lambda c: f"{c}_MW")
        row.index = [m]
        row["Total_MW"] = fuel.sum()
        trow = pd.DataFrame([by_tech])
        trow.index = [m]
        monthly = row.combine_first(monthly) if not monthly.empty else row
        tech = trow.combine_first(tech) if not tech.empty else trow
        cols = FUEL_COLS
        for r in REGIONS:
            rr = region_row(by_region[r], m)
            regional[r] = rr.combine_first(regional[r]) if not regional[r].empty else rr
            regional[r] = regional[r].reindex(columns=[c for c in cols if c in regional[r]] + ["Total_MW"]).sort_index()
        monthly = monthly.reindex(columns=[c for c in cols if c in monthly] + ["Total_MW"]).sort_index()
        tech = tech[tech.iloc[-1].sort_values(ascending=False).index].sort_index()
        save(args.out, monthly, tech, regional, with_cf=m == todo[-1])
        done += 1
    print(f"Saved {args.out}: {done} months fetched, {len(monthly)} months in total", flush=True)
    if monthly.empty:
        raise SystemExit("No capacity data")


if __name__ == "__main__":
    main()
