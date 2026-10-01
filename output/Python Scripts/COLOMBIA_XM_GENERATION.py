"""
Colombia generation by fuel from XM (the Colombian market operator)'s
free, no-key public API (servapibi.xm.com.co) - the Colombia equivalent
of Brazil's generation mix, with gas split out (Colombia's thermal fleet
burns domestic gas and, since 2016, LNG through Cartagena).

Two XM calls, both confirmed against live responses (Sep-2026, via
COLOMBIA_XM_GENERATION_DISCOVERY.py):
  - /lists  MetricId=ListadoRecursos, Entity=Sistema: every plant's code
    and its fuel ('EnerSource': AGUA, GAS, CARBON, RAD SOLAR, VIENTO,
    COMBUSTOLEO, ACPM, BAGAZO, ...).
  - /hourly MetricId=Gene, Entity=Recurso: each plant's hourly output in
    kWh ('Hour01'..'Hour24'). The /daily endpoint rejects this metric.
    XM publishes with a lag of about two weeks.

Each plant's day is summed and grouped by its fuel. The per-fuel daily
totals are kept in co_generation_daily_by_fuel.csv, so later runs only
ask XM for the last few weeks rather than the whole history.

Outputs:
  colombia_generation_daily.xlsx  'Data'     daily mean GW by category
                                  'By fuel'  daily MWh per XM fuel label
  co_generation_daily_by_fuel.csv  the cache (MWh per fuel label per day)

Usage: python3 COLOMBIA_XM_GENERATION.py [--full]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

LISTS_URL = "https://servapibi.xm.com.co/lists"
HOURLY_URL = "https://servapibi.xm.com.co/hourly"
START_DATE = dt.date(2020, 1, 1)
CHUNK_DAYS = 14   # ~500 plants x 24 hours a day - keeps each response a few MB
REFRESH_DAYS = 21  # XM fills in the latest ~2 weeks late and revises them
CACHE_CSV = "co_generation_daily_by_fuel.csv"
OUT_FILE = "colombia_generation_daily.xlsx"
HEADERS = {"Connection": "close"}

# XM fuel label ('EnerSource') -> category
FUEL_TO_CATEGORY = {
    "AGUA": "hydro",
    "GAS": "gas", "GLP": "gas",
    "CARBON": "coal",
    "COMBUSTOLEO": "liquids", "ACPM": "liquids", "JET-A1": "liquids",
    "VIENTO": "wind",
    "RAD SOLAR": "solar",
    "BAGAZO": "biomass", "BIOGAS": "biomass", "BIOMASA": "biomass",
}
CATEGORIES = ["hydro", "gas", "coal", "liquids", "wind", "solar", "biomass", "unclassified"]


def post(url, body, tries=3):
    for attempt in range(tries):
        try:
            r = requests.post(url, json=body, headers=HEADERS, timeout=180)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            if attempt == tries - 1:
                raise
            print(f"  retrying after {type(e).__name__}", flush=True)
            time.sleep(5 * (attempt + 1))


def plant_fuels():
    """Plant code -> XM fuel label."""
    payload = post(LISTS_URL, {"MetricId": "ListadoRecursos", "Entity": "Sistema"})
    fuels = {}
    for item in payload.get("Items", []):
        for entity in item.get("ListEntities", []):
            v = entity.get("Values", {})
            if v.get("Code"):
                fuels[v["Code"]] = str(v.get("EnerSource") or "").strip().upper() or "UNKNOWN"
    print(f"  {len(fuels):,} plants listed", flush=True)
    return fuels


def fetch_generation(start, end, fuels):
    """Daily MWh per fuel label, start..end inclusive."""
    rows = []
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=CHUNK_DAYS - 1), end)
        try:
            payload = post(HOURLY_URL, {"MetricId": "Gene", "Entity": "Recurso",
                                        "StartDate": cur.isoformat(), "EndDate": chunk_end.isoformat()})
        except requests.RequestException as e:
            print(f"  {cur}..{chunk_end}: FAILED ({type(e).__name__}: {e})", flush=True)
            cur = chunk_end + dt.timedelta(days=1)
            continue
        totals = {}
        for item in payload.get("Items", []):
            day = item.get("Date")
            for entity in item.get("HourlyEntities", []):
                v = entity.get("Values", {})
                kwh = sum(float(x) for k, x in v.items() if k.startswith("Hour") and x not in (None, ""))
                fuel = fuels.get(v.get("code"), "UNKNOWN")
                totals[(day, fuel)] = totals.get((day, fuel), 0.0) + kwh / 1000  # kWh -> MWh
        rows += [{"Date": d, "fuel": f, "MWh": m} for (d, f), m in totals.items()]
        days_back = len({d for d, _ in totals})
        print(f"  {cur}..{chunk_end}: {days_back} day(s)", flush=True)
        cur = chunk_end + dt.timedelta(days=1)
        time.sleep(0.3)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).pivot_table(index="Date", columns="fuel", values="MWh", aggfunc="sum")
    df.index = pd.to_datetime(df.index)
    return df


def to_categories(by_fuel):
    unknown = [c for c in by_fuel.columns if c not in FUEL_TO_CATEGORY]
    if unknown:
        print(f"  fuel labels counted as 'unclassified': {unknown}", flush=True)
    out = pd.DataFrame(index=by_fuel.index)
    for cat in CATEGORIES:
        cols = [c for c in by_fuel.columns if FUEL_TO_CATEGORY.get(c, "unclassified") == cat]
        out[cat] = by_fuel[cols].sum(axis=1) if cols else 0.0
    out = out.loc[:, (out != 0).any()]  # drop categories XM never reports
    return (out / 24).round(1)  # MWh a day -> mean MW


NOTES_LINES = [
    "UNITS",
    "'Data': MW - each day's mean generation by category (daily MWh / 24). "
    "'By fuel': MWh generated that day per XM fuel label.",
    "",
    "CATEGORIES",
    "hydro = AGUA; gas = GAS, GLP; coal = CARBON; liquids = COMBUSTOLEO, ACPM, JET-A1; wind = VIENTO; "
    "solar = RAD SOLAR; biomass = BAGAZO, BIOGAS, BIOMASA; unclassified = plants missing from XM's list.",
    "",
    "SOURCE",
    "XM (servapibi.xm.com.co): hourly generation per plant (MetricId Gene, Entity Recurso, kWh) summed by "
    "each plant's fuel from XM's plant list (ListadoRecursos). XM publishes about two weeks behind, so the "
    "latest days fill in on later runs.",
]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--full", action="store_true", help="ignore the cache and fetch everything from START_DATE")
    args = parser.parse_args()

    cached = None
    if not args.full and os.path.exists(CACHE_CSV):
        cached = pd.read_csv(CACHE_CSV, index_col=0, parse_dates=True)
    start = START_DATE if cached is None or cached.empty else \
        max(START_DATE, (cached.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date())
    end = dt.date.today()
    print(f"Plant list...", flush=True)
    fuels = plant_fuels()
    print(f"Hourly generation {start}..{end}...", flush=True)
    fresh = fetch_generation(start, end, fuels)
    if cached is not None and not cached.empty:
        by_fuel = fresh.combine_first(cached) if not fresh.empty else cached
    else:
        by_fuel = fresh
    if by_fuel.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    by_fuel = by_fuel.sort_index().fillna(0.0).round(1)
    by_fuel.index.name = "Date"
    by_fuel.to_csv(CACHE_CSV)

    data = to_categories(by_fuel)
    xlsx_notes.write_workbook(OUT_FILE, {"Data": data, "By fuel": by_fuel}, NOTES_LINES,
                              {"UNITS", "CATEGORIES", "SOURCE"})
    print(f"\nSaved {OUT_FILE} ({len(data):,} days, {data.index.min():%d-%b-%Y} to {data.index.max():%d-%b-%Y})")
    print(data.tail(7).to_string())


if __name__ == "__main__":
    main()
