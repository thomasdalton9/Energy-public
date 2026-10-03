"""
Sri Lanka electricity generation by source from PUCSL's GenData platform (Public Utilities
Commission of Sri Lanka, data from the CEB system control centre), https://gendata.pucsl.gov.lk/.
Found via discovery_archive/subcontinent/SEA_DISCOVERY_SOUTH_ASIA.py (earlier probes in
discovery_archive/rest_of_world/SRILANKA_* got HTTP 500 because the API needs its query parameters).

API (no key):
  /api/actual-system-dispatch?dateAggregation=15min&from=<ISO>&to=<ISO>   15-minute dispatch per plant,
        MW ({reportTimestamp, powerPlantId, dispatchValueInMW}); at most 7 days per call; from 2023-01-01
  /api/metadata/power-plants   plant list with technology, capacity and energy type (Major hydro, Oil (CEB),
        Coal, Wind, Solar, Mini hydro, Biomass ...), used to map plants to fuels

Writes output/Data and Chart Outputs/sri_lanka_power_generation_daily.xlsx:
  Daily        standard layout, MWh per day by fuel (mean of the day's 15-minute MW x 24 per plant)
  Demand       daily peak and average of total dispatch, MW
  By plant     MWh per day per plant (latest 120 days)
  Plants       the plant list with the fuel each is counted under

Timestamps are treated as Sri Lanka local time (the API labels them ...Z but its days run midnight to
midnight local; see the solar-profile check printed each run). Dispatch covers plants dispatched by the
system control centre; small rooftop solar is not in it.

Incremental: the Daily sheet is the history store; only missing days (plus REVISION_DAYS) are fetched.
Runs on the 1st and 15th.

    python3 asia/SRI_LANKA_PUCSL.py
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

API = "https://gendata.pucsl.gov.lk/api"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"}
T = (15, 180)
DATA_START = date(2023, 1, 1)
CHUNK_DAYS = 7
REVISION_DAYS = 3
CHECKPOINT = 25          # write the workbook every 25 chunks during a backfill
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "sri_lanka_power_generation_daily.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Other"]
# keyword in the plant's energy type / technology / fuel (lower case) -> standard fuel; first match wins
FUEL_KEYS = [("bess", "Other"), ("battery", "Other"), ("hydro", "Hydro"), ("coal", "Coal"), ("wind", "Wind"), ("solar", "Solar"), ("biomass", "Bioenergy"),
             ("dendro", "Bioenergy"), ("bio", "Bioenergy"), ("lng", "Gas"), ("natural gas", "Gas"), ("oil", "Oil"),
             ("diesel", "Oil"), ("naphtha", "Oil"), ("fuel", "Oil"), ("thermal", "Oil"), ("combined cycle", "Oil")]


def out(*a):
    print(*a, flush=True)


def get(path, **params):
    for i in range(4):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=T)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                raise
            out(f"  retry {path} {params}: {e}")
            time.sleep(5 * (i + 1))


def plants():
    rows = []
    for p in get("metadata/power-plants").get("data", []):
        cx = p.get("powerPlantComplex") or {}
        et = cx.get("energyType") or {}
        text = " ".join(str(x) for x in (et.get("name"), et.get("slug"), p.get("technology"), p.get("fuelUsed"),
                                         cx.get("name"), p.get("name"))).lower()
        fuel = next((f for k, f in FUEL_KEYS if k in text), "Other")
        rows.append({"id": p["id"], "name": p.get("name"), "complex": cx.get("name"), "energy_type": et.get("name"),
                     "technology": p.get("technology"), "fuel_used": p.get("fuelUsed"),
                     "capacity_MW": p.get("capacityData"), "fuel": fuel})
    df = pd.DataFrame(rows).drop_duplicates("id", keep="first")   # the list repeats some plant ids
    return df.set_index("id")


def fetch(d0, d1):
    """15-minute dispatch for local days d0..d1 (inclusive) -> long frame (time, plant, MW)."""
    j = get("actual-system-dispatch", dateAggregation="15min",
            **{"from": datetime.combine(d0, datetime.min.time()).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
               "to": datetime.combine(d1 + timedelta(days=1), datetime.min.time()).strftime("%Y-%m-%dT%H:%M:%S.000Z")})
    df = pd.DataFrame(j.get("data") or [])
    if df.empty:
        return df
    df["time"] = pd.to_datetime(df["reportTimestamp"]).dt.tz_localize(None)
    df["MW"] = pd.to_numeric(df["dispatchValueInMW"], errors="coerce")
    return df[["time", "powerPlantId", "MW"]]


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def save(frames, old, old_dem, old_plant, meta, out_path):
    """Aggregate the fetched 15-minute rows, merge with the saved history and write the workbook (also used as a
    checkpoint during long backfills)."""
    gen, dem, by_plant = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    if frames:
        raw = pd.concat(frames, ignore_index=True)
        raw["date"] = raw["time"].dt.normalize()
        # per plant: mean MW over the day x 24 (bridges missing intervals)
        mwh = raw.groupby(["date", "powerPlantId"])["MW"].mean().mul(24).unstack()
        by_plant = mwh.rename(columns=lambda c: meta["name"].get(c, f"plant {c}"))
        fuel_of = {c: meta["fuel"].get(c, "Other") for c in mwh.columns}
        gen = mwh.T.groupby(fuel_of).sum().T
        gen = gen.reindex(columns=FUELS, fill_value=0.0).add_suffix("_MWh")
        gen.columns.name = by_plant.columns.name = None
        gen["Total_MWh"] = gen.sum(axis=1)
        gen["Intervals"] = raw.groupby("date")["time"].nunique()
        gen = gen.round(1)
        gen = gen[gen["Total_MWh"] > 0]   # days not published yet come back as all-zero rows: fetch them next run
        tot = raw.groupby("time")["MW"].sum()
        g = tot.groupby(tot.index.normalize())
        dem = pd.DataFrame({"Demand_avg_MW": g.mean().round(0), "Demand_peak_MW": g.max().round(0),
                            "Demand_min_MW": g.min().round(0)})
        dem = dem[dem.index.isin(gen.index)]
        solar = [c for c, f in fuel_of.items() if f == "Solar"]
        if solar:
            sol = raw[raw["powerPlantId"].isin(solar)]
            prof = sol.groupby(sol["time"].dt.hour)["MW"].mean().round(0)
            out(f"  solar profile by hour (peak should be ~12-13 if timestamps are local): {prof.to_dict()}")

    daily = merge(old, gen)
    demand = merge(old_dem, dem)
    plant_daily = merge(old_plant, by_plant)
    if daily.empty:
        raise SystemExit("No PUCSL dispatch data")
    plant_daily = plant_daily[plant_daily.index >= plant_daily.index.max() - pd.Timedelta(days=120)]
    for f in (daily, demand, plant_daily):
        f.index.name = "date"
    notes = [
        "UNITS",
        "Daily: MWh per day by fuel = for each plant, mean of the day's 15-minute dispatch (MW) x 24, summed by fuel. "
        "Intervals = 15-minute timestamps returned for the day (96 = complete).",
        "Demand: total dispatch of all plants, MW - daily average, peak and minimum (a proxy for system demand "
        "met by dispatched plants).",
        "By plant: MWh per day per plant, last 120 days (the full per-plant history is not kept).",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. Plants dispatched by the CEB system "
        "control centre (major hydro, CEB and IPP thermal, Lakvijaya coal, wind and solar farms, mini hydro where "
        "reported). Rooftop solar is not included. Fuel per plant from the PUCSL plant list (energy type / "
        "technology): Hydro includes mini hydro; Oil covers CEB and IPP oil-fired plants (Sri Lanka has no gas "
        "supply - the combined-cycle plants burn oil); see the Plants sheet.",
        "",
        "SOURCE",
        "PUCSL GenData (Public Utilities Commission of Sri Lanka), actual system dispatch: "
        "https://gendata.pucsl.gov.lk/ (API /api/actual-system-dispatch, /api/metadata/power-plants).",
    ]
    xlsx_notes.write_workbook(out_path, {"Daily": daily, "Demand": demand, "By plant": plant_daily,
                                         "Plants": meta.reset_index()}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {out_path}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", default=DATA_START.isoformat())
    args = ap.parse_args()
    start = date.fromisoformat(args.start)

    meta = plants()
    out(f"{len(meta)} plants; by fuel: {meta['fuel'].value_counts().to_dict()}")
    other = meta[meta["fuel"] == "Other"]
    if not other.empty:
        out(f"  unmapped (counted as Other): {other[['name', 'energy_type', 'technology']].to_dict('records')}")

    old, old_dem, old_plant = (read_sheet(args.out, s) for s in ("Daily", "Demand", "By plant"))
    if not old.empty:   # all-zero days were saved before they were published: fetch them again
        old = old[old["Total_MWh"] > 0]
        old_dem = old_dem[old_dem.index.isin(old.index)] if not old_dem.empty else old_dem
    yesterday = date.today() - timedelta(days=1)
    have = set(old.index.date) if not old.empty else set()
    revise = {yesterday - timedelta(days=k) for k in range(REVISION_DAYS)}
    todo = [d for d in (start + timedelta(days=k) for k in range((yesterday - start).days + 1))
            if d not in have or d in revise]
    out(f"{len(have)} days saved; fetching {len(todo)}")

    frames = []
    i = 0
    while i < len(todo):
        d0 = todo[i]
        d1 = min(d0 + timedelta(days=CHUNK_DAYS - 1), yesterday)
        df = fetch(d0, d1)
        if not df.empty:
            frames.append(df[df["time"].dt.date.isin([d for d in todo if d0 <= d <= d1])])
        out(f"  {d0}..{d1}: {len(df)} rows")
        if len(frames) % CHECKPOINT == 0 and frames:
            save(frames, old, old_dem, old_plant, meta, args.out)   # checkpoint: a timeout keeps what is done
        while i < len(todo) and todo[i] <= d1:
            i += 1
        time.sleep(0.5)

    save(frames, old, old_dem, old_plant, meta, args.out)


if __name__ == "__main__":
    main()
