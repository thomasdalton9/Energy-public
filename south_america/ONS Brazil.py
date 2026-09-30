"""
Brazil generation by source (hydro, thermal, wind, solar) from ONS open
data - one script for both the daily and the monthly view.

Source: ONS's hourly energy balance by subsystem
(BALANCO_ENERGIA_SUBSISTEMA, one parquet file per year). Everything
else is built from those hourly readings:

  brazil_generation_daily.xlsx    'Data'      national daily mean MW by source
                                              (what south_america_dashboard.py reads)
                                  'By region' the same per ONS subsystem
  brazil_generation_monthly.xlsx  'National'  monthly mean MW and GWh by source
                                  'By region' monthly mean GW, peak GW and GWh
                                              per subsystem and source

Cache: each year's file is kept in ons_cache/ next to this script. Past
years are downloaded once; this year's (and last year's, until April,
while ONS may still revise it) is re-downloaded only when ONS's copy has
changed (its ETag). So a normal run fetches one file, not seventeen.
--refresh re-downloads everything; --cache-dir puts the cache elsewhere.

Usage: python3 "ONS Brazil.py" [--refresh] [--cache-dir DIR]
"""

import argparse
import io
import json
import os
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes/api_keys

import xlsx_notes

HERE = Path(__file__).resolve().parent
START_YEAR = 2010
URL = "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/balanco_energia_subsistema_ho/BALANCO_ENERGIA_SUBSISTEMA_{year}.parquet"
COLS = {"val_gerhidraulica": "hydro", "val_gertermica": "thermal", "val_gereolica": "wind", "val_gersolar": "solar"}
SOURCES = list(COLS.values())
# ONS subsystem codes -> names; any whole-system row (SIN) is left out of
# the national sum so it isn't counted twice
REGIONS = {"N": "North", "NE": "Northeast", "S": "South", "SE": "Southeast/Centre-West", "SE/CO": "Southeast/Centre-West"}
WHOLE_SYSTEM = {"SIN"}
DAILY_OUT = HERE / "brazil_generation_daily.xlsx"
MONTHLY_OUT = HERE / "brazil_generation_monthly.xlsx"


# ------------------------------------------------------------------
# CACHE
# ------------------------------------------------------------------

def is_settled(year, today):
    """Past years stop changing; last year's file gets until April for
    ONS's late revisions."""
    return year < today.year - 1 or (year == today.year - 1 and today.month > 3)


def fetch_year(year, cache_dir, state, refresh, today):
    """The year's hourly balance, from the cache when it's still current.
    Returns None if ONS hasn't published the year (yet)."""
    path = cache_dir / f"BALANCO_ENERGIA_SUBSISTEMA_{year}.parquet"
    url = URL.format(year=year)
    if path.exists() and not refresh and is_settled(year, today):
        print(f"  {year}: cached")
        return pd.read_parquet(path)
    try:
        head = requests.head(url, timeout=60)
        if head.status_code in (403, 404):
            print(f"  {year}: not published")
            return None
        head.raise_for_status()
        etag = head.headers.get("ETag")
        if path.exists() and not refresh and etag and etag == state.get(str(year)):
            print(f"  {year}: cached (unchanged at ONS)")
            return pd.read_parquet(path)
        print(f"  {year}: downloading...")
        r = requests.get(url, timeout=600)
        r.raise_for_status()
        df = pd.read_parquet(io.BytesIO(r.content))
    except requests.RequestException as e:
        if path.exists():
            print(f"  {year}: download failed ({type(e).__name__}) - using the cached copy")
            return pd.read_parquet(path)
        raise
    tmp = path.with_suffix(".tmp")
    tmp.write_bytes(r.content)
    tmp.replace(path)
    state[str(year)] = r.headers.get("ETag") or etag
    return df


def load_hourly(cache_dir, refresh):
    cache_dir.mkdir(parents=True, exist_ok=True)
    state_path = cache_dir / "etags.json"
    state = json.loads(state_path.read_text()) if state_path.exists() else {}
    today = date.today()
    frames = []
    try:
        for year in range(START_YEAR, today.year + 1):
            df = fetch_year(year, cache_dir, state, refresh, today)
            if df is not None and len(df):
                frames.append(df)
    finally:
        state_path.write_text(json.dumps(state, indent=1))
    if not frames:
        raise RuntimeError("No ONS data downloaded or cached")
    return pd.concat(frames, ignore_index=True)


# ------------------------------------------------------------------
# SHAPE
# ------------------------------------------------------------------

def tidy(raw):
    """Hourly MW per subsystem and source: time, region, hydro..solar."""
    sub_col = "id_subsistema" if "id_subsistema" in raw.columns else "nom_subsistema"
    df = pd.DataFrame({"time": pd.to_datetime(raw["din_instante"]), "code": raw[sub_col].astype(str).str.strip().str.upper()})
    for col, name in COLS.items():
        df[name] = pd.to_numeric(raw[col], errors="coerce") if col in raw.columns else float("nan")
    codes = sorted(df["code"].unique())
    print(f"  subsystems in the file: {codes}")
    unknown = [c for c in codes if c not in REGIONS and c not in WHOLE_SYSTEM]
    if unknown:
        print(f"  WARNING: unrecognised subsystem codes {unknown} - kept as their own regions")
    df = df[~df["code"].isin(WHOLE_SYSTEM)].copy()
    df["region"] = df["code"].map(REGIONS).fillna(df["code"])
    # one row per hour and region (SE and SE/CO are the same subsystem)
    return df.groupby(["time", "region"], as_index=False)[SOURCES].sum(min_count=1)


def drop_incomplete_tail(hourly):
    """ONS's latest day is often still filling in (a source missing for
    some regions or hours), which would read as a collapse. Drop trailing
    days where any source has fewer readings than a normal full day."""
    days = hourly["time"].dt.date
    counts = hourly[SOURCES].notna().groupby(days).sum()
    recent = counts.tail(30)
    full = recent.max()
    active = full > 0
    complete = (counts.loc[:, active] >= full[active]).all(axis=1)
    last_complete = complete[complete].index.max()
    dropped = [d for d in counts.index if d > last_complete]
    if dropped:
        print(f"  dropping incomplete latest day(s): {', '.join(f'{d:%d-%b-%Y}' for d in dropped)}")
    return hourly[days <= last_complete]


def national_hourly(hourly):
    return hourly.groupby("time")[SOURCES].sum(min_count=1)


def daily_tables(hourly):
    nat = national_hourly(hourly)
    daily = nat.groupby(nat.index.date).mean()
    daily.index.name = "date"
    reg = hourly.assign(date=hourly["time"].dt.date).groupby(["date", "region"])[SOURCES].mean().unstack("region")
    reg = reg.swaplevel(axis=1)
    reg = reg[[(region, source) for region in sorted(reg.columns.levels[0]) for source in SOURCES
               if (region, source) in reg.columns]]
    reg.columns = [f"{region} {source}" for region, source in reg.columns]
    return daily.round(1), reg.round(1)


def hours_per_reading(times):
    """Hours each reading stands for, per month - 1 for hourly data, 0.5
    if ONS ever posts half-hourly - so GWh doesn't depend on it."""
    t = pd.Series(sorted(set(times)))
    step = t.diff().dt.total_seconds().div(3600)
    return step.groupby(t.dt.to_period("M")).median().fillna(1.0)


def monthly_tables(hourly):
    nat = national_hourly(hourly)
    month = nat.index.to_period("M")
    step = hours_per_reading(nat.index)
    mean_mw = nat.groupby(month).mean()
    hours = nat.groupby(month).size() * step.reindex(mean_mw.index).fillna(1.0)
    national = pd.DataFrame(index=mean_mw.index)
    for s in SOURCES:
        national[f"{s}_avg_mw"] = mean_mw[s].round(1)
    for s in SOURCES:
        national[f"{s}_gwh"] = (mean_mw[s] * hours / 1000).round(1)
    national["total_gwh"] = national[[f"{s}_gwh" for s in SOURCES]].sum(axis=1, min_count=1)
    national["hours_of_data"] = hours
    national.index = national.index.to_timestamp()
    national.index.name = "month"

    long = hourly.melt(id_vars=["time", "region"], value_vars=SOURCES, var_name="source", value_name="mw")
    long["month"] = long["time"].dt.to_period("M").dt.to_timestamp()
    reg = long.groupby(["month", "region", "source"]).agg(avg_mw=("mw", "mean"), peak_mw=("mw", "max"),
                                                          hours=("mw", "count")).reset_index()
    reg["hours"] = reg["hours"] * reg["month"].dt.to_period("M").map(step).fillna(1.0).values
    reg["generation_gwh"] = (reg["avg_mw"] * reg["hours"] / 1000).round(1)
    reg["average_generation_gw"] = (reg["avg_mw"] / 1000).round(3)
    reg["peak_generation_gw"] = (reg["peak_mw"] / 1000).round(3)
    reg = reg.dropna(subset=["avg_mw"]).set_index("month")
    return national, reg[["region", "source", "generation_gwh", "average_generation_gw", "peak_generation_gw", "hours"]]


# ------------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------------

SOURCE_NOTE = ["SOURCE",
               "ONS (Operador Nacional do Sistema Eletrico) open data - BALANCO_ENERGIA_SUBSISTEMA hourly parquet "
               "files, one per year, from ons-aws-prod-opendata.s3.amazonaws.com. Columns val_gerhidraulica / "
               "val_gertermica / val_gereolica / val_gersolar.",
               "",
               "REGIONS",
               "ONS's four subsystems: North, Northeast, South, Southeast/Centre-West. National = their sum "
               "(any whole-system 'SIN' row is left out so nothing is counted twice)."]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--refresh", action="store_true", help="re-download every year, ignoring the cache")
    parser.add_argument("--cache-dir", default=str(HERE / "ons_cache"), help="where the yearly files are kept")
    args = parser.parse_args()

    print("ONS energy balance by year:")
    hourly = drop_incomplete_tail(tidy(load_hourly(Path(args.cache_dir), args.refresh)))
    print(f"  {len(hourly):,} hourly region rows, {hourly['time'].min():%d-%b-%Y} to {hourly['time'].max():%d-%b-%Y %H:%M}")

    daily, daily_reg = daily_tables(hourly)
    xlsx_notes.write_workbook(
        DAILY_OUT, {"Data": daily, "By region": daily_reg},
        ["UNITS",
         "MW - each day's mean of the hourly readings ('average MW for that day'), not daily energy. "
         "Daily energy in MWh = MW x 24.",
         "",
         "TABS",
         "'Data': national generation by source (the four subsystems summed). "
         "'By region': the same per subsystem, one column per region and source.",
         ""] + SOURCE_NOTE,
        {"UNITS", "TABS", "SOURCE", "REGIONS"})

    national, monthly_reg = monthly_tables(hourly)
    xlsx_notes.write_workbook(
        MONTHLY_OUT, {"National": national, "By region": monthly_reg},
        ["UNITS",
         "*_avg_mw: the month's mean of the hourly readings (MW). *_gwh: energy generated in the month "
         "(mean MW x hours of data / 1000). The latest month is partial - see hours_of_data.",
         "'By region': average and peak GW and GWh per subsystem and source; 'hours' = readings behind each row.",
         ""] + SOURCE_NOTE,
        {"UNITS", "SOURCE", "REGIONS"})

    print(f"\nSaved {DAILY_OUT.name} (Data, By region) and {MONTHLY_OUT.name} (National, By region)")
    print("\nLatest days (national, MW):")
    print(daily.tail().to_string())
    print("\nLatest months (national GWh):")
    print(national[[f"{s}_gwh" for s in SOURCES] + ["total_gwh", "hours_of_data"]].tail(4).to_string())


if __name__ == "__main__":
    main()
