"""
Shared helpers for the Europe national-TSO pulls (GB NESO, Belgium Elia, Denmark Energinet, Spain REE, ...).

Each pull produces a standard raw generation workbook "<country>_power_generation_daily.xlsx":
sheet "Daily" with date (index) and Hydro/Gas/Wind/Solar/Coal/Nuclear/Oil/Bioenergy/Other/Total _MWh columns
(MWh per UTC day), the layout add_charts.power_daily() and the master workbooks expect.

The committed workbook is the history store: a run reads it and fetches only days not saved yet plus a short
revision window (TSOs restate recent days).
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

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
FUEL_COLS = ["Hydro_MWh", "Gas_MWh", "Wind_MWh", "Solar_MWh", "Coal_MWh", "Nuclear_MWh", "Oil_MWh",
             "Bioenergy_MWh", "Other_MWh"]
DEFAULT_START = date(2021, 1, 1)
RELOAD_DAYS = 14
NOTES_TITLES = {"UNITS", "COLUMNS", "COVERAGE", "SOURCE"}


def get_json(url, params=None, attempts=4, timeout=(15, 120), headers=None):
    """GET -> parsed JSON with retries (the TSO open-data APIs are public and occasionally slow)."""
    last = None
    for i in range(1, attempts + 1):
        try:
            r = requests.get(url, params=params, headers=headers or HEADERS, timeout=timeout)
            if r.status_code >= 400:
                raise requests.HTTPError(f"{r.status_code}: {r.text[:300]}")
            return r.json()
        except (requests.RequestException, ValueError) as e:
            last = e
            print(f"    attempt {i}/{attempts} failed: {type(e).__name__}: {str(e)[:300]}", file=sys.stderr, flush=True)
            if i < attempts:
                time.sleep(5 * i)
    raise last


def bucket(name):
    """TSO fuel / technology label -> one of the dashboard's columns, or None for storage (it shifts energy, it is
    not generation)."""
    n = str(name).lower()
    if any(k in n for k in ("storage", "pump", "battery", "import", "interconnect")):
        return None
    if "bio" in n:
        return "Bioenergy_MWh"
    if "nuclear" in n:
        return "Nuclear_MWh"
    if "wind" in n:
        return "Wind_MWh"
    if "solar" in n or "photovolt" in n:
        return "Solar_MWh"
    if "hydro" in n or "water" in n or "run-of" in n or "run of" in n:
        return "Hydro_MWh"
    if "coal" in n or "lignite" in n or "anthracite" in n:
        return "Coal_MWh"
    if "gas" in n or "combined cycle" in n:
        return "Gas_MWh"
    if any(k in n for k in ("oil", "diesel", "fuel")):
        return "Oil_MWh"
    return "Other_MWh"


def to_daily(df, value_cols_to_bucket=None):
    """Frame indexed by date with FUEL_COLS (any missing -> 0) -> add Total_MWh, round."""
    out = df.reindex(columns=FUEL_COLS).fillna(0.0)
    out["Total_MWh"] = out[FUEL_COLS].sum(axis=1)
    out.index = pd.to_datetime(out.index).date
    out.index.name = "date"
    return out.round(1)


def load_archive(path):
    try:
        d = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index).date
    d.index.name = "date"
    return d


def run(out_name, fetch, notes_lines, description=""):
    """Standard incremental run. fetch(start, end) -> to_daily() frame for the complete UTC days start..end."""
    ap = argparse.ArgumentParser(description=description)
    ap.add_argument("--out", default=os.path.join(DATA_DIR, out_name))
    ap.add_argument("--start-date", type=date.fromisoformat, default=DEFAULT_START)
    args = ap.parse_args()
    existing = load_archive(args.out)
    end = date.today() - timedelta(days=1)
    start = args.start_date if existing.empty else max(args.start_date, max(existing.index) - timedelta(days=RELOAD_DAYS))
    print(f"{out_name}: fetching {start} to {end} ...", file=sys.stderr, flush=True)
    new = fetch(start, end)
    if new is None or new.empty:
        if existing.empty:
            raise RuntimeError("no data returned and no archive exists")
        print("no new rows; archive unchanged", file=sys.stderr)
        return
    combined = new if existing.empty else pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": combined}, notes_lines, NOTES_TITLES)
    print(combined.tail(5).to_string(), file=sys.stderr)
    print(f"Saved {args.out} ({len(combined)} days, {combined.index.min()} to {combined.index.max()})")


def month_chunks(start, end):
    """(first day, last day) pairs covering start..end by calendar month."""
    cur = date(start.year, start.month, 1)
    while cur <= end:
        nxt = date(cur.year + (cur.month == 12), cur.month % 12 + 1, 1)
        yield max(cur, start), min(nxt - timedelta(days=1), end)
        cur = nxt


STANDARD_NOTES = [
    "UNITS",
    "MWh of electricity generated per UTC day. Storage (pumped hydro, batteries) is left out: it shifts energy "
    "rather than generating it. Total_MWh is the sum of the fuel columns.",
    "",
    "COLUMNS",
    "Hydro, Gas, Wind, Solar, Coal, Nuclear, Oil (-> 'Other Fossil' on the charts), Bioenergy, Other.",
    "",
]
