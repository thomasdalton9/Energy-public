"""
US natural gas from EIA: monthly consumption by sector, dry production
and trade (Natural Gas Monthly), and weekly working gas in storage
(Weekly Natural Gas Storage Report). Also writes Mexico's gas imports
from the US (US pipeline + LNG exports to Mexico) - Mexico's main gas
supply - to a separate workbook for the North America master.

  https://api.eia.gov/v2/seriesid/NG.<series>  (EIA API v2, needs EIA_API_KEY)
  fallback, no key: https://www.eia.gov/dnav/ng/hist_xls/<series><m|w>.xls

Incremental: each run re-reads only the last 24 months (monthly - EIA
revises recent months for about two years) / 8 weeks (storage) and
merges them into the saved sheets; a new file backfills from 2015 (the
storage water-year chart needs five prior years).

Usage: python3 US_GAS_EIA.py [--out "output/Data and Chart Outputs/us_gas.xlsx"]
                             [--mexico-out "output/Data and Chart Outputs/mexico_gas.xlsx"]
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

API = "https://api.eia.gov/v2/seriesid/NG.{sid}.{freq}"
XLS = "https://www.eia.gov/dnav/ng/hist_xls/{sid}{freq}.xls"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 90)
HISTORY_START = "2015-01-01"
REFRESH_MONTHS = 24
REFRESH_WEEKS = 8
DEFAULT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")

# column -> EIA series (MMcf per month); converted to Bcf/d
DEMAND = {
    "Residential": "N3010US2",
    "Commercial": "N3020US2",
    "Industrial": "N3035US2",
    "Electric_power": "N3045US2",
    "Vehicle_fuel": "N3025US2",
    "Lease_and_plant_fuel": "N9160US2",
    "Pipeline_and_distribution": "N9170US2",
    "Total_consumption": "N9140US2",
}
SUPPLY = {
    "Dry_production": "N9070US2",
    "Pipeline_imports_from_Canada": "N9102CN2",
    "LNG_imports": "N9103US2",
    "Total_imports": "N9100US2",
    "LNG_exports": "N9133US2",
    "Pipeline_exports_to_Mexico": "N9132MX2",
    "Pipeline_exports_to_Canada": "N9132CN2",
    "Total_exports": "N9130US2",
}
MEXICO = {
    "Pipeline_imports_from_US": "N9132MX2",   # = US pipeline exports to Mexico
    "LNG_imports_from_US": "N9133MX2",        # = US LNG exports to Mexico (trucked/shipped, small)
}
# Weekly working gas in underground storage, Bcf (EIA's five storage regions since 2015)
STORAGE = {
    "Lower_48": "NW2_EPG0_SWO_R48_BCF",
    "East": "NW2_EPG0_SWO_R31_BCF",
    "Midwest": "NW2_EPG0_SWO_R32_BCF",
    "South_Central": "NW2_EPG0_SWO_R33_BCF",
    "Mountain": "NW2_EPG0_SWO_R34_BCF",
    "Pacific": "NW2_EPG0_SWO_R35_BCF",
}


def get(url, **kw):
    last = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"    attempt {attempt + 1}/3 failed: {type(e).__name__}: {str(e)[:200]}", flush=True)
            time.sleep(5 * (attempt + 1))
    raise last


def fetch_api(sid, freq, start):
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise RuntimeError("EIA_API_KEY not set")
    params = {"api_key": key, "start": start, "length": 5000}
    rows = get(API.format(sid=sid, freq=freq.upper()), params=params).json()["response"]["data"]
    # The seriesid route can return more than one row per period (other units such as $/Mcf, or other
    # breakdowns of the same flow). Keep only this series' volume rows; never let a stray row overwrite one.
    kinds = sorted({(str(r.get("series", "")), str(r.get("units", "")), str(r.get("process-name", r.get("process", ""))))
                    for r in rows})
    if len(kinds) > 1:
        print(f"    {sid}: API returned {len(kinds)} row kinds {kinds} - keeping {sid} volume rows only", flush=True)
    rows = [r for r in rows
            if str(r.get("series", sid)).upper().endswith(sid.upper())
            and re.search(r"MMCF|BCF", str(r.get("units", "MMCF")), re.I)]
    by_period = {}
    for r in rows:
        by_period.setdefault(r["period"], []).append(r["value"])
    dupes = {p: v for p, v in by_period.items() if len(set(map(str, v))) > 1}
    if dupes:
        print(f"    {sid}: {len(dupes)} period(s) with conflicting values, left blank: {dict(list(dupes.items())[:5])}",
              flush=True)
    return pd.Series({p: v[0] for p, v in by_period.items() if p not in dupes}, dtype="object")


def fetch_xls(sid, freq, start):
    raw = pd.read_excel(io.BytesIO(get(XLS.format(sid=sid, freq=freq.lower())).content), sheet_name="Data 1", header=2)
    s = pd.Series(raw.iloc[:, 1].values, index=raw.iloc[:, 0].values)
    return s[pd.to_datetime(s.index, errors="coerce") >= pd.Timestamp(start)]


def fetch(sid, freq, start):
    """One EIA series from `start` as a float Series indexed by date (month start or week ending)."""
    for name, fn in (("API", fetch_api), ("XLS", fetch_xls)):
        try:
            s = fn(sid, freq, start)
        except Exception as e:  # noqa: BLE001 - try the next source
            print(f"  {sid} via {name} failed: {type(e).__name__}: {str(e)[:200]}", flush=True)
            continue
        s.index = pd.to_datetime(pd.Index(s.index).astype(str), errors="coerce")
        s = pd.to_numeric(s, errors="coerce")
        s = s[s.index.notna()].dropna().sort_index()
        if freq == "m":
            s.index = s.index.to_period("M").to_timestamp()
        s = s[s.index >= pd.Timestamp(start)]   # the seriesid route returns full history whatever `start` says
        if not s.empty:
            print(f"  {sid} via {name}: {len(s)} rows {s.index.min():%Y-%m-%d}..{s.index.max():%Y-%m-%d}", flush=True)
            return s
    return pd.Series(dtype=float, index=pd.DatetimeIndex([]))


def load(path, sheet, col):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d[col] = pd.to_datetime(d[col].astype(str), errors="coerce")
    return d.dropna(subset=[col]).set_index(col).sort_index()


def start_for(saved, freq):
    if saved.empty:
        return HISTORY_START
    off = pd.DateOffset(months=REFRESH_MONTHS) if freq == "m" else pd.DateOffset(weeks=REFRESH_WEEKS)
    return max(pd.Timestamp(HISTORY_START), saved.index.max() - off).strftime("%Y-%m-%d")


def pull(series, saved, freq, bcfd=True):
    """Fetch every series from the refresh start, convert MMcf/month to Bcf/d, merge over the saved rows."""
    suffix = "_Bcf_per_day" if bcfd else "_Bcf"
    if not saved.empty:   # columns retired by a series change are dropped, not carried forward
        saved = saved[[c for c in saved.columns if c[:-len(suffix)] in series and c.endswith(suffix)]]
    start = start_for(saved, freq)
    new = pd.DataFrame({c: fetch(sid, freq, start) for c, sid in series.items()})
    new.index = pd.DatetimeIndex(new.index)
    failed = [c for c in series if c not in new or new[c].isna().all()]
    if new.empty:
        return saved, failed
    if bcfd:
        new = new.div(new.index.days_in_month, axis=0) / 1000.0   # MMcf/month -> Bcf/d
        new = new.add_suffix("_Bcf_per_day")
    else:
        new = new.add_suffix("_Bcf")
    out = new.combine_first(saved) if not saved.empty else new
    out = out[[c for c in new.columns] + [c for c in out.columns if c not in new.columns]].sort_index()
    return out.round(3), failed


def frame(d, name):
    """Saved layout: date index (written as the first column), months as YYYY-MM."""
    d = d.copy()
    if name == "Month":
        d.index = d.index.strftime("%Y-%m")
    d.index.name = name
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DEFAULT_DIR, "us_gas.xlsx"))
    ap.add_argument("--mexico-out", default=os.path.join(DEFAULT_DIR, "mexico_gas.xlsx"))
    args = ap.parse_args()

    print("Demand by sector (monthly)", flush=True)
    demand, f1 = pull(DEMAND, load(args.out, "Demand by sector", "Month"), "m")
    print("Supply and trade (monthly)", flush=True)
    supply, f2 = pull(SUPPLY, load(args.out, "Supply and trade", "Month"), "m")
    print("Working gas in storage (weekly)", flush=True)
    storage, f3 = pull(STORAGE, load(args.out, "Storage weekly", "date"), "w", bcfd=False)
    print("Mexico imports from the US (monthly)", flush=True)
    mexico, f4 = pull(MEXICO, load(args.mexico_out, "Imports from US", "Month"), "m")
    failed = f1 + f2 + f3
    if failed:
        print(f"WARNING: no data this run for {failed}", flush=True)
    if demand.empty and supply.empty and storage.empty:
        raise SystemExit("No EIA data fetched")

    def cover(d):
        return f"{d.index.min():%b %Y} to {d.index.max():%b %Y}" if not d.empty else "none"

    notes = [
        "UNITS",
        "Demand by sector, Supply and trade: billion cubic feet per day (Bcf/d), monthly average = EIA monthly "
        "volume (MMcf) / days in month / 1,000.",
        "Storage weekly: working gas in underground storage, Bcf, week ending Friday (EIA Weekly Natural Gas "
        "Storage Report). Lower_48 plus EIA's five regions.",
        "Lease_and_plant_fuel, Pipeline_and_distribution: gas used in field operations and processing plants, and "
        "by pipeline compressors and distribution systems; Total_consumption includes them.",
        "",
        "COVERAGE",
        f"Demand: {cover(demand)}. Supply and trade: {cover(supply)}. Storage: {cover(storage)}. EIA's monthly "
        "data run about two months behind (Natural Gas Monthly); recent months are revised, so each run "
        f"re-reads the last {REFRESH_MONTHS} months.",
        "",
        "SOURCE",
        "US Energy Information Administration (EIA), API v2 (series IDs: " +
        ", ".join(f"{c}={s}" for c, s in {**DEMAND, **SUPPLY, **STORAGE}.items()) +
        "). Fallback: EIA's keyless history workbooks https://www.eia.gov/dnav/ng/hist_xls/<series>.xls",
        "https://www.eia.gov/naturalgas/data.php",
    ]
    sheets = {"Demand by sector": frame(demand, "Month"), "Supply and trade": frame(supply, "Month"),
              "Storage weekly": frame(storage, "date")}
    sheets = {k: v for k, v in sheets.items() if len(v)}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: " + ", ".join(f"{k} {len(v)} rows" for k, v in sheets.items()), flush=True)

    if not mexico.empty:
        mx_notes = [
            "UNITS",
            "Billion cubic feet per day (Bcf/d), monthly average. Mexico imports most of its gas by pipeline from "
            "the US; these are EIA's US export volumes to Mexico (pipeline, and LNG by truck/ship), so they are "
            "Mexico's imports from the US as measured on the US side.",
            "LNG_imports_from_US is EIA's series as published, but it has months of 2-3.5 Bcf/d (20-25% of all US LNG "
            "exports) among months of ~0.1 Bcf/d, which is not credible; it is kept here but not charted. Pipeline "
            "imports are the meaningful series (they match EIA's US pipeline exports to Mexico).",
            "",
            "COVERAGE",
            f"{cover(mexico)}. Mexico's own demand-by-sector statistics (SENER SIE) are not pulled yet.",
            "",
            "SOURCE",
            "EIA API v2: N9132MX2 (US natural gas pipeline exports to Mexico), N9133MX2 (US LNG exports to Mexico).",
            "https://www.eia.gov/dnav/ng/ng_move_expc_s1_m.htm",
        ]
        xlsx_notes.write_workbook(args.mexico_out, {"Imports from US": frame(mexico, "Month")}, mx_notes,
                                  {"UNITS", "COVERAGE", "SOURCE"})
        print(f"Saved {args.mexico_out}: {len(mexico)} rows", flush=True)


if __name__ == "__main__":
    main()
