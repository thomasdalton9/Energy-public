"""
New Zealand power generation by fuel, daily, from the Electricity Authority's
EMI metered generation dataset (half-hourly kWh per generating unit, grid
and embedded plants the Authority reconciles):

  https://emidatasets.blob.core.windows.net/publicdata/Datasets/Wholesale/Generation/Generation_MD/YYYYMM_Generation_MD.csv

Writes nz_power_generation_daily.xlsx, sheet "Daily" in the standard layout
(date, <Fuel>_MWh, Total_MWh): Hydro, Gas, Wind, Solar, Coal, Oil (diesel),
Bioenergy (wood, biogas), Other (geothermal and anything unmapped) - the
EMI fuel code totals are kept on sheet "By fuel code".

Incremental: re-reads the latest two saved months (the Authority revises
the current month's file) plus any newer monthly files; a new file
backfills from 2021. Saves after every month.

Usage: python3 NZ_EMI_GENERATION.py [--out "output/Data and Chart Outputs/nz_power_generation_daily.xlsx"]
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

BLOB = "https://emidatasets.blob.core.windows.net/publicdata"
PREFIX = "Datasets/Wholesale/Generation/Generation_MD/"
HISTORY_START = "2021-01"
REFRESH_MONTHS = 2
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "nz_power_generation_daily.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Other"]
# EMI Fuel_Code -> standard fuel (unlisted codes -> Other, and are printed so they can be added)
FUEL = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal", "Diesel": "Oil",
        "Wood": "Bioenergy", "Biogas": "Bioenergy", "Geo": "Other", "Geothermal": "Other"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url, **kw):
    for attempt in range(4):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 180), **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    attempt {attempt + 1}/4 failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(f"failed: {url}")


def monthly_files():
    """{month Timestamp: blob name} for every YYYYMM_Generation_MD.csv in the dataset folder."""
    r = get(BLOB, params={"restype": "container", "comp": "list", "prefix": PREFIX, "maxresults": 5000})
    out = {}
    for name in re.findall(r"<Name>([^<]+)</Name>", r.text):
        m = re.search(r"/(\d{6})_Generation_MD\.csv$", name)
        if m:
            out[pd.Timestamp(m.group(1)[:4] + "-" + m.group(1)[4:] + "-01")] = name
    return out


def fetch_month(name):
    """One monthly file -> daily MWh by EMI fuel code."""
    d = pd.read_csv(io.BytesIO(get(f"{BLOB}/{name}").content), low_memory=False)
    tps = [c for c in d.columns if re.fullmatch(r"TP\d+", str(c))]   # TP1..TP50 (46/50 on daylight-saving days)
    d["kWh"] = d[tps].apply(pd.to_numeric, errors="coerce").sum(axis=1)
    d["date"] = pd.to_datetime(d["Trading_Date"], errors="coerce")
    d = d.dropna(subset=["date"])
    by_code = d.pivot_table(index="date", columns="Fuel_Code", values="kWh", aggfunc="sum") / 1000.0   # kWh -> MWh
    return by_code


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def standard(by_code):
    unknown = sorted(set(by_code.columns) - set(FUEL))
    if unknown:
        print(f"  NOTE fuel codes counted as Other: {unknown}", flush=True)
    g = by_code.T.groupby(lambda c: FUEL.get(c, "Other")).sum(min_count=1).T
    g = g.reindex(columns=[f for f in FUELS if f in g.columns]).add_suffix("_MWh")
    g["Total_MWh"] = by_code.sum(axis=1, min_count=1)
    return g


def save(path, codes):
    daily = standard(codes)
    notes = [
        "UNITS",
        "MWh per day (sum of the half-hourly metered kWh / 1,000). Daily: Hydro, Gas, Wind, Solar, Coal, Oil (EMI "
        "fuel code Diesel), Bioenergy (Wood, Biogas), Other (Geo = geothermal, ~18% of NZ generation, plus any "
        "unmapped code). Huntly's Rankine units are reported under the fuel code EMI gives them (Coal or Gas).",
        "By fuel code: MWh per day by the EMI Fuel_Code as published.",
        "",
        "COVERAGE",
        f"New Zealand, daily, {daily.index.min():%d %b %Y} to {daily.index.max():%d %b %Y}. Generation the "
        "Electricity Authority reconciles (grid-connected and larger embedded plants); small rooftop solar is not "
        "metered here. The current month's file is published during the month and revised.",
        "",
        "SOURCE",
        f"Electricity Authority (Te Mana Hiko), EMI Generation_MD dataset: {BLOB}/{PREFIX}",
        "https://www.emi.ea.govt.nz/Wholesale/Datasets/Generation/Generation_MD",
    ]
    c = codes.copy()
    for x in (daily, c):
        x.index.name = "date"
    xlsx_notes.write_workbook(path, {"Daily": daily.round(1), "By fuel code": c.round(1)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-months", type=int, default=100)
    args = ap.parse_args()

    codes = load(args.out, "By fuel code")
    files = {m: n for m, n in monthly_files().items() if m >= pd.Timestamp(HISTORY_START + "-01")}
    if not files:
        raise SystemExit("No Generation_MD files listed")
    have = sorted(set(codes.index.to_period("M").to_timestamp())) if not codes.empty else []
    todo = [m for m in files if m not in have] + [m for m in have[-REFRESH_MONTHS:] if m in files]
    todo = sorted(set(todo))[:args.max_months]
    print(f"{len(files)} monthly files {min(files):%Y-%m}..{max(files):%Y-%m}; {len(have)} months saved; "
          f"fetching {len(todo)}", flush=True)
    for m in todo:
        new = fetch_month(files[m])
        if new.empty:
            print(f"  {m:%Y-%m}: no rows", flush=True)
            continue
        if not codes.empty:   # replace the whole month (a revised file replaces, never adds to, the saved one)
            codes = codes[codes.index.to_period("M") != m.to_period("M")]
        codes = pd.concat([codes, new]).sort_index()
        print(f"  {m:%Y-%m}: {len(new)} days, {new.sum().sum() / 1000:.0f} GWh", flush=True)
        save(args.out, codes)
    if codes.empty:
        raise SystemExit("No NZ generation data")
    print(f"Saved {args.out}: {len(codes)} days", flush=True)


if __name__ == "__main__":
    main()
