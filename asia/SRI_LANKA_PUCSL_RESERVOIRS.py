"""
Sri Lanka major hydro reservoir storage from PUCSL's GenData platform (Public Utilities Commission of Sri Lanka,
data from the CEB system control centre), https://gendata.pucsl.gov.lk/ (dashboard "Reservoir Data > Reservoir
Storage Level"). Found via discovery_archive/asia/HYDRO_SSEA_DISCOVERY1-3.py (the endpoint is in the site's JS bundle).

API (no key): /api/reservoir/storage-rainfall?dateAggregation=day&from=<ISO>&to=<ISO>
  -> one row per reservoir per day: reportDate, reservoirName, storageInGwh (energy stored, GWh), rainfallInMm (catchment
     rainfall, mm), molBelowSpillInM (the minimum operating level, metres below spill - a constant per reservoir).
  Daily from 2013-01-01. Six CEB storage reservoirs: Castlereigh and Maussakele (Laxapana cascade), Kotmale, Victoria,
  Randenigala (Mahaweli complex) and Samanalawewa.

Writes output/Data and Chart Outputs/sri_lanka_hydro_reservoirs.xlsx:
  Daily      date x reservoir: storage in GWh (<Reservoir>_GWh) and Total_storage_GWh (sum of the six)
  Rainfall   date x reservoir: catchment rainfall, mm
  Limits     per reservoir: minimum operating level, metres below spill
  Water year chart (Oct-Sep) of total storage, via water_year_chart (add_charts.py)

Incremental: the Daily sheet is the history store; only years with missing days are fetched (one call per year), plus
the last REVISION_DAYS. Runs on the 1st and 15th.

    python3 asia/SRI_LANKA_PUCSL_RESERVOIRS.py
"""
import argparse
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

API = "https://gendata.pucsl.gov.lk/api/reservoir/storage-rainfall"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"}
T = (15, 180)
DATA_START = date(2013, 1, 1)
REVISION_DAYS = 18   # runs are 14-17 days apart: re-read everything since the last run, plus spare
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "sri_lanka_hydro_reservoirs.xlsx")
RES = ["Castlereigh", "Maussakele", "Kotmale", "Victoria", "Randenigala", "Samanalawewa"]
ALIAS = {"maussakalle": "Maussakele", "maussakelle": "Maussakele", "samanala wewa": "Samanalawewa",
         "samanalawewa": "Samanalawewa", "castlereagh": "Castlereigh"}


def out(*a):
    print(*a, flush=True)


def name(x):
    k = str(x).strip().lower()
    return ALIAS.get(k) or next((r for r in RES if r.lower() == k), str(x).strip())


def fetch(d0, d1):
    """Rows for d0..d1 inclusive."""
    p = {"dateAggregation": "day", "from": f"{d0}T00:00:00.000Z", "to": f"{d1 + timedelta(days=1)}T00:00:00.000Z"}
    for i in range(4):
        try:
            r = requests.get(API, params=p, headers=H, timeout=T)
            r.raise_for_status()
            d = pd.DataFrame(r.json().get("data") or [])
            if d.empty:
                return d
            d["date"] = pd.to_datetime(d["reportDate"].str[:10])
            d["res"] = d["reservoirName"].map(name)
            return d[(d.date.dt.date >= d0) & (d.date.dt.date <= d1)]
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                out(f"  {d0}..{d1} failed: {e}")
                return None
            time.sleep(5 * (i + 1))


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    if old.empty:
        return new
    if new.empty:
        return old
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    old, old_rain = read_sheet(args.out, "Daily"), read_sheet(args.out, "Rainfall")
    today = date.today()
    have = set(old.index.date) if not old.empty else set()
    want = [DATA_START + timedelta(days=k) for k in range((today - DATA_START).days + 1)]
    # days the feed never had stay missing: only the last year's gaps (and the revision window) are asked for again
    recent = today - timedelta(days=365) if have else DATA_START
    todo = [d for d in want if (d not in have and d >= recent) or d >= today - timedelta(days=REVISION_DAYS)]
    years = sorted({d.year for d in todo})
    out(f"{len(have)} days saved; {len(todo)} to fetch in years {years}")
    parts = []
    for y in years:
        d0, d1 = max(date(y, 1, 1), DATA_START), min(date(y, 12, 31), today)
        d = fetch(d0, d1)
        if d is not None and not d.empty:
            parts.append(d)
            out(f"  {y}: {len(d)} rows, {d.date.nunique()} days")
    if not parts:
        if old.empty:
            raise SystemExit("No PUCSL reservoir data")
        out("Nothing new")
        return
    raw = pd.concat(parts)
    other = raw.loc[~raw.res.isin(RES), "res"].value_counts()
    if len(other):   # the feed's own 'Total' row and small reservoirs it carried in some years
        out(f"Left out (not one of the six): {other.to_dict()}")
    raw = raw[raw.res.isin(RES)]
    st = raw.pivot_table(index="date", columns="res", values="storageInGwh", aggfunc="last")
    rain = raw.pivot_table(index="date", columns="res", values="rainfallInMm", aggfunc="last")
    st = st[(st.fillna(0) > 0).any(axis=1)]   # days the feed carries as all-zero are missing, not empty reservoirs
    cols = [r for r in RES if r in st] + [c for c in st if c not in RES]
    st = st[cols].rename(columns=lambda c: f"{c.replace(' ', '_')}_GWh")
    st["Total_storage_GWh"] = st.sum(axis=1, min_count=len(cols)).round(1)
    rain = rain.reindex(st.index)[cols].rename(columns=lambda c: f"{c.replace(' ', '_')}_mm")
    daily, rainfall = merge(old, st), merge(old_rain, rain)
    daily.index.name = rainfall.index.name = "date"
    lim = (raw.sort_values("date").groupby("res")["molBelowSpillInM"].last().reindex(cols).rename("MOL_below_spill_m")
           .to_frame())
    lim.index.name = "reservoir"
    notes = [
        "UNITS",
        "Daily: energy stored in each reservoir, GWh (the energy its water would generate through the downstream "
        "cascade, as CEB reports it), at the morning reading; Total_storage_GWh = sum of the six (blank if one is "
        "missing). Rainfall: catchment rainfall, mm per day. Limits: minimum operating level (MOL), metres below spill.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. CEB's six major storage reservoirs: "
        "Castlereigh and Maussakele (Laxapana / Kelani cascade), Kotmale, Victoria and Randenigala (Mahaweli complex), "
        "Samanalawewa (Walawe). Days the feed reports as all zero are left out.",
        "",
        "SOURCE",
        "Public Utilities Commission of Sri Lanka (PUCSL) GenData, Reservoir Storage Level (data from the CEB system "
        "control centre): https://gendata.pucsl.gov.lk/reservoir-storage-level "
        "(API https://gendata.pucsl.gov.lk/api/reservoir/storage-rainfall).",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Rainfall": rainfall, "Limits": lim}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
