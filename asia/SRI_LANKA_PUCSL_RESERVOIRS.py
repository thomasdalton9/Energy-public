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
  Raw        storage per reservoir exactly as the API returns it (the history store the checks run on)
  Flags      reservoir-days blanked in Daily by the check below

Data check: the API's 'day' aggregation (the only one it serves: 15min / hour / raw return nothing, checked in
discovery_archive/asia/QAFIX_DISCOVERY1.py) SUMS the readings of a day when the database holds more than one, so
some days come back 2-3x (2014-12-04: Victoria 1,294.8 vs ~430; 10-19 Mar 2015 every reservoir doubled), and in
Nov 2015 Randenigala alternates between ~4 and ~113 GWh. Storage moves slowly, so a value is blanked when it is
more than JUMP_REL (and JUMP_ABS GWh) away from the last accepted value both reading forwards and reading backwards
in time (a genuine step, e.g. cyclone inflows, passes one of the two directions; a spike or a wrong block fails
both). A day with FLAG_DAY or more reservoirs flagged is blanked for all six (a summed day).
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
JUMP_REL, JUMP_ABS, REANCHOR, FLAG_DAY = 0.6, 10.0, 30, 3
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


def one_pass(s):
    """Reject values more than JUMP_REL x (and JUMP_ABS GWh) away from the last accepted value; after REANCHOR
    rejections in a row the current value is accepted as the new level."""
    rej, last, streak = pd.Series(False, index=s.index), None, 0
    for t, v in s.items():
        if pd.isna(v):
            continue
        if last is None or abs(v - last) <= max(JUMP_REL * last, JUMP_ABS) or streak >= REANCHOR:
            last, streak = v, 0
        else:
            rej[t], streak = True, streak + 1
    return rej


def check(raw):
    """Raw storage (date x <Reservoir>_GWh) -> (clean storage, flags frame)."""
    cols = [c for c in raw.columns if c.endswith("_GWh")]
    bad = pd.DataFrame({c: one_pass(raw[c]) & one_pass(raw[c][::-1])[::-1] for c in cols}, index=raw.index)
    whole = bad.sum(axis=1) >= FLAG_DAY
    bad.loc[whole] = raw.loc[whole, cols].notna().values
    clean = raw[cols].mask(bad)
    flags = [{"date": t, "reservoir": c.replace("_GWh", ""), "value_GWh": raw.at[t, c],
              "reason": "summed day (several reservoirs out of line)" if whole[t] else "out of line both ways"}
             for t in raw.index[bad.any(axis=1)] for c in cols if bad.at[t, c]]
    flags = pd.DataFrame(flags, columns=["date", "reservoir", "value_GWh", "reason"]).set_index("date")
    return clean, flags


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
    old, old_rain = read_sheet(args.out, "Raw"), read_sheet(args.out, "Rainfall")
    if old.empty:   # first run with the check: the saved Daily sheet is still the data as published
        old = read_sheet(args.out, "Daily").drop(columns=["Total_storage_GWh"], errors="ignore")
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
    if not parts and old.empty:
        raise SystemExit("No PUCSL reservoir data")
    if parts:
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
        rain = rain.reindex(st.index)[cols].rename(columns=lambda c: f"{c.replace(' ', '_')}_mm")
        lim = (raw.sort_values("date").groupby("res")["molBelowSpillInM"].last().reindex(cols)
               .rename("MOL_below_spill_m").to_frame())
        lim.index.name = "reservoir"
    else:   # nothing fetched: the saved history is still re-checked and rewritten
        out("Nothing new")
        st, rain = pd.DataFrame(), pd.DataFrame()
        cols = [c for c in old.columns if c.endswith("_GWh")]
        try:
            lim = pd.read_excel(args.out, sheet_name="Limits", index_col=0)
        except (FileNotFoundError, ValueError):
            lim = pd.DataFrame()
    raw_st, rainfall = merge(old, st), merge(old_rain, rain)
    daily, flags = check(raw_st)
    daily["Total_storage_GWh"] = daily.sum(axis=1, min_count=len(cols)).round(1)
    out(f"check: {len(flags)} reservoir-days blanked on {flags.index.nunique()} days; total storage max "
        f"{raw_st.sum(axis=1, min_count=len(cols)).max():,.0f} -> {daily['Total_storage_GWh'].max():,.0f} GWh")
    daily.index.name = rainfall.index.name = raw_st.index.name = "date"
    notes = [
        "UNITS",
        "Daily: energy stored in each reservoir, GWh (the energy its water would generate through the downstream "
        "cascade, as CEB reports it), at the morning reading; Total_storage_GWh = sum of the six (blank if one is "
        "missing). Rainfall: catchment rainfall, mm per day (as published). Limits: minimum operating level (MOL), "
        "metres below spill. Raw: storage exactly as the API returns it. Flags: values blanked in Daily.",
        "",
        "VALIDATION",
        "The API's daily figure SUMS a day's readings when the database holds more than one, so some days come back "
        "2-3x (2014-12-04: Victoria 1,294.8 vs ~430 GWh, total 3,278 vs ~1,100; every reservoir doubled 10-19 Mar "
        "2015), and in Nov 2015 Randenigala alternates between ~4 and ~113 GWh. No other aggregation is served, so "
        f"these cannot be de-duplicated at source: a value more than {JUMP_REL:.0%} (and {JUMP_ABS:.0f} GWh) away "
        "from the last accepted value reading both forwards and backwards in time is blanked, and a day with "
        f"{FLAG_DAY} or more reservoirs flagged is blanked for all (Flags sheet). Genuine fast changes (e.g. the "
        "Nov 2025 cyclone inflows) pass. The total is blank on any day with a blanked reservoir.",
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
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Rainfall": rainfall, "Limits": lim, "Raw": raw_st,
                                         "Flags": flags}, notes, {"UNITS", "COVERAGE", "VALIDATION", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
