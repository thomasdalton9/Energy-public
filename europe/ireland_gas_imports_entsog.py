"""
Daily gas flows into (and through) the Republic of Ireland from the
ENTSOG Transparency Platform - the up-to-yesterday complement to GNI's
open-data supply file (ireland_gas_supply.py), which is only
republished quarterly with a ~4-5 month lag (confirmed ending
2026-03-31 as of 2026-09-30).

Points (from IRELAND_ENTSOG_DISCOVERY.py; GNI is operator IE-TSO-0002):
  ITP-00495  Moffat (IE)       entry  gas arriving from GB into the ROI system
  ITP-00222  South North CSEP  exit   gas leaving the ROI system onward to Northern Ireland
ENTSOG lists no production point for Corrib or Inch, so indigenous
production stays GNI-only; here Net_Imports_ROI_GWh = Moffat entry
minus the South-North exit, which is what should line up with GNI's
"Moffat ROI" column over the overlap (the run log prints that check).

Indicator "Physical Flow", periodType=day. ENTSOG reports kWh/d;
converted to GWh/d to match the GNI files. Maintains a growing archive:
every run re-pulls the last RELOAD_DAYS (ENTSOG restates recent days)
and upserts.
"""

import argparse
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

API = "https://transparency.entsog.eu/api/v1/operationalData"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
TIMEOUT = (15, 120)
FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [10, 30]

POINTS = {
    "Moffat_Entry_GWh": {"pointKey": "ITP-00495", "operatorKey": "IE-TSO-0002", "directionKey": "entry"},
    "SouthNorth_Exit_to_NI_GWh": {"pointKey": "ITP-00222", "operatorKey": "IE-TSO-0002", "directionKey": "exit"},
}
UNIT_TO_GWH = {"kWh/d": 1e-6, "MWh/d": 1e-3, "GWh/d": 1.0}

DATA_START = date(2015, 1, 1)
CHUNK_DAYS = 366
RELOAD_DAYS = 45

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "ireland_gas_imports_entsog_daily.xlsx")
GNI_SUPPLY = os.path.join(REPO_ROOT, "output", "ireland_gas_supply_daily.xlsx")


def fetch(point, start, end):
    params = {"periodType": "day", "indicator": "Physical Flow", "from": start.isoformat(), "to": end.isoformat(),
              "limit": -1, **point}
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(API, headers=HEADERS, params=params, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json().get("operationalData", [])
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr, flush=True)
            if attempt < FETCH_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS[attempt - 1])
    raise last_error


def fetch_range(start, end):
    frames = {}
    for col, point in POINTS.items():
        rows = []
        chunk_start = start
        while chunk_start <= end:
            chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS - 1), end)
            print(f"  {col}: {chunk_start} to {chunk_end} ...", file=sys.stderr, flush=True)
            rows.extend(fetch(point, chunk_start, chunk_end))
            chunk_start = chunk_end + timedelta(days=1)
        if not rows:
            print(f"  {col}: no rows", file=sys.stderr, flush=True)
            continue
        df = pd.DataFrame(rows)
        units = set(df["unit"].dropna())
        unknown = units - set(UNIT_TO_GWH)
        if unknown:
            raise RuntimeError(f"{col}: unexpected unit(s) {unknown} - sample row: {rows[0]}")
        df["gwh"] = pd.to_numeric(df["value"], errors="coerce") * df["unit"].map(UNIT_TO_GWH)
        # periodFrom carries mixed +00:00/+01:00 offsets across DST; the gas
        # day is what the local date of periodFrom names.
        df["date"] = pd.to_datetime(df["periodFrom"], utc=True).dt.tz_convert("Europe/Dublin").dt.date
        s = df.groupby("date")["gwh"].sum()
        print(f"  {col}: {len(s)} days, {s.index.min()} to {s.index.max()}, mean {s.mean():.1f} GWh/d, "
              f"units seen {sorted(units)}", file=sys.stderr, flush=True)
        frames[col] = s
    if not frames:
        return pd.DataFrame()
    wide = pd.DataFrame(frames).sort_index()
    wide.index.name = "date"
    if {"Moffat_Entry_GWh", "SouthNorth_Exit_to_NI_GWh"} <= set(wide.columns):
        wide["Net_Imports_ROI_GWh"] = wide["Moffat_Entry_GWh"] - wide["SouthNorth_Exit_to_NI_GWh"].fillna(0)
    return wide


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def compare_with_gni(archive):
    try:
        gni = pd.read_excel(GNI_SUPPLY, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return
    gni.index = pd.to_datetime(gni.index).date
    both = archive.join(gni[["Moffat_Imports_GWh"]], how="inner").dropna()
    if both.empty:
        return
    last = both.iloc[-365:]
    for col in ["Moffat_Entry_GWh", "Net_Imports_ROI_GWh"]:
        if col in last:
            diff = (last[col] - last["Moffat_Imports_GWh"])
            print(f"  vs GNI 'Moffat ROI' (last {len(last)} overlapping days): {col} mean diff {diff.mean():+.1f} GWh/d, "
                  f"mean abs diff {diff.abs().mean():.1f} GWh/d", file=sys.stderr, flush=True)


NOTES_LINES = [
    "UNITS",
    "GWh of gas per gas day (ENTSOG publishes kWh/d; divided by 1,000,000).",
    "",
    "COLUMNS",
    "Moffat_Entry_GWh: physical flow entering GNI's Republic of Ireland system at Moffat (IE), i.e. all gas "
    "arriving from Great Britain, including gas that transits onward to Northern Ireland.",
    "SouthNorth_Exit_to_NI_GWh: physical flow leaving the ROI system at the South North CSEP, i.e. onward "
    "to Northern Ireland.",
    "Net_Imports_ROI_GWh: Moffat entry minus South North exit - the gas retained for the Republic of "
    "Ireland, comparable to the 'Moffat ROI' column in ireland_gas_supply_daily.xlsx (GNI's own file). "
    "See the run log for the measured difference over the overlap.",
    "",
    "COVERAGE",
    f"Daily from {DATA_START.isoformat()} (or whenever ENTSOG's series begins) to yesterday; the last "
    f"{RELOAD_DAYS} days are re-pulled every run because ENTSOG restates recent days. Indigenous production "
    "(Corrib) is NOT on ENTSOG - GNI's quarterly open-data file remains the only source for it.",
    "",
    "SOURCE",
    f"ENTSOG Transparency Platform operational data API: {API} (indicator 'Physical Flow', periodType day).",
]
NOTES_SECTION_TITLES = {"UNITS", "COLUMNS", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    existing = load_archive(args.out)
    stop_at = date.today() - timedelta(days=1)
    if existing.empty:
        start = args.start_date
    else:
        start = max(args.start_date, max(existing.index) - timedelta(days=RELOAD_DAYS))
    print(f"Fetching {start} to {stop_at} ...", file=sys.stderr, flush=True)
    new = fetch_range(start, stop_at)

    if new.empty:
        combined = existing
    elif existing.empty:
        combined = new
    else:
        combined = pd.concat([existing, new])
        combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    combined.index.name = "date"
    if combined.empty:
        raise RuntimeError("No data at all - ENTSOG returned nothing and no archive exists")

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(combined.tail(7).round(1).to_string(), file=sys.stderr, flush=True)
    compare_with_gni(combined)
    print(f"Saved to {args.out} ({len(combined)} days, {combined.index.min()} to {combined.index.max()})")


if __name__ == "__main__":
    main()
