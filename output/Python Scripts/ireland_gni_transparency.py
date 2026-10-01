"""
Ireland daily gas ENTRY FLOWS by entry point and gas CONSUMPTION by
market sector, from Gas Networks Ireland's data-transparency pages -
the up-to-date source (through yesterday), unlike GNI's open-data CSVs
on data.gov.ie which are republished quarterly and currently stop at
2026-03-31.

Pages (confirmed via IRELAND_GNI_TRANSPARENCY_DISCOVERY.py):
  /about/data-transparency/entry-flows/physical-flows
  /about/data-transparency/exit-flows/gas-consumption-by-market-sector
Both are Highcharts widgets fed by /api/v1/{physicalflows,gasconsumption}
(JSON, last few periods only) and an "Export all data" button that hits
  https://www.gasnetworks.ie/csv/{dataset}?frequency=daily&date=A&date_end=B
returning CSV rows "Name,Location,Date,Value,Unit" (kWh). This script
uses the CSV export with a date range, chunked month by month, and
keeps two archives in one workbook (GWh/d):
  Entry flows             one column per entry point (Bellanaboy=Corrib, Moffat, ...) + Aggregate
  Consumption by sector   one column per market sector (DM, NDM, ...) - the
                          "Forecast EOD" series are dropped, only outturn kept
Incremental: each run re-pulls the last RELOAD_DAYS and upserts.
"""

import argparse
import io
import os
import re
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

CSV_URL = "https://www.gasnetworks.ie/csv/{dataset}"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/csv,*/*",
    "Referer": "https://www.gasnetworks.ie/about/data-transparency/",
}
TIMEOUT = (15, 120)
FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [10, 30]

DATASETS = {
    "physicalflows": {"sheet": "Entry flows", "keep_group": None, "drop_group": None},
    "gasconsumption": {"sheet": "Consumption by sector", "keep_group": None, "drop_group": re.compile("forecast", re.I)},
}
# Only needed from where GNI's quarterly open-data files stop (they run
# 2018-01-01 to 2026-03-31); the one-day overlap is a sanity check.
DATA_START = date(2026, 3, 31)
CHUNK_DAYS = 31
RELOAD_DAYS = 45
KWH_TO_GWH = 1e-6

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "ireland_gni_transparency_daily.xlsx")


def fetch_csv(dataset, start, end):
    # GNI's export treats date_end as EXCLUSIVE (confirmed live: every
    # chunk came back one day short), so ask for one day past `end`.
    params = {"frequency": "daily", "date": start.isoformat(), "date_end": (end + timedelta(days=1)).isoformat()}
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(CSV_URL.format(dataset=dataset), headers=HEADERS, params=params, timeout=TIMEOUT)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr, flush=True)
            if attempt < FETCH_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS[attempt - 1])
    raise last_error


def column_name(location):
    return re.sub(r"[^A-Za-z0-9]+", "_", str(location)).strip("_") + "_GWh"


def fetch_range(dataset, start, end, drop_group):
    frames = []
    combos = set()
    chunk_start = start
    while chunk_start <= end:
        chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS - 1), end)
        print(f"  {dataset}: {chunk_start} to {chunk_end} ...", file=sys.stderr, flush=True)
        text = fetch_csv(dataset, chunk_start, chunk_end)
        chunk_start = chunk_end + timedelta(days=1)
        if not text.strip() or "Value" not in text.splitlines()[0]:
            print(f"    unexpected/empty response: {text[:200]!r}", file=sys.stderr, flush=True)
            continue
        df = pd.read_csv(io.StringIO(text))
        if df.empty:
            print("    0 rows", file=sys.stderr, flush=True)
            continue
        df.columns = [c.strip() for c in df.columns]
        df["group"] = df["Name"].str.split(" - ", n=1).str[1].str.strip()
        combos |= set(zip(df["group"], df["Location"]))
        if drop_group is not None:
            df = df[~df["group"].astype(str).str.contains(drop_group)]
        units = set(df["Unit"].dropna().astype(str).str.strip())
        if units - {"kWh"}:
            raise RuntimeError(f"{dataset}: unexpected unit(s) {units}")
        df["date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
        df = df.dropna(subset=["date"])
        df["gwh"] = pd.to_numeric(df["Value"], errors="coerce") * KWH_TO_GWH
        got = df["date"].max() if not df.empty else None
        print(f"    {len(df)} rows, last day {got}", file=sys.stderr, flush=True)
        frames.append(df[["date", "Location", "gwh"]])
    print(f"  {dataset}: series seen (group, location): {sorted(combos)}", file=sys.stderr, flush=True)
    if not frames:
        return pd.DataFrame()
    long = pd.concat(frames)
    wide = long.pivot_table(index="date", columns="Location", values="gwh", aggfunc="sum")
    wide.columns = [column_name(c) for c in wide.columns]
    wide.index.name = "date"
    wide = wide.sort_index()
    # GNI's per-point series can have blank days while Aggregate is complete
    # (seen live: 21 blank Moffat days in the first 177) - surface it.
    gaps = wide.isna().sum()
    if gaps.any():
        print(f"  {dataset}: blank days per column: {gaps[gaps > 0].to_dict()}", file=sys.stderr, flush=True)
    return wide


def load_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new):
    if new.empty:
        return existing
    if existing.empty:
        return new
    combined = pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    combined.index.name = "date"
    return combined


NOTES_LINES = [
    "UNITS",
    "All values in GWh per gas day (GNI publishes kWh; divided by 1,000,000).",
    "",
    "ENTRY FLOWS",
    "One column per entry point as GNI names them: Bellanaboy is the Corrib field's onshore terminal "
    "(indigenous production, = 'Corrib' in the open-data supply file); Moffat is the WHOLE interconnector "
    "from Great Britain, including gas that transits the ROI system onward to Northern Ireland - subtract "
    "(Total_LDM - ROI_LDM) from the consumption sheet to get the ROI-only figure that matches the "
    "open-data 'Moffat ROI' column (reconciled exactly on the 2026-03-31 overlap day); Gormanston is the "
    "South-North entry (normally 0); Aggregate is GNI's total of all entry points.",
    "",
    "CONSUMPTION BY SECTOR",
    "As GNI names them: NDM (non-daily metered residential/small commercial), DM (daily metered "
    "industrial), ROI_Power_Gen (gas-fired power stations), ROI_LDM (large daily metered INCLUDING power "
    "generation - subtract ROI_Power_Gen for the open-data 'LDM non Power Gen' figure), Total_LDM (ROI_LDM "
    "plus the Northern Ireland/Isle of Man offtake). NDM + DM + ROI_LDM = the open-data 'Total ROI demand'. "
    "Only outturn 'Gas Consumption' series are kept; GNI's 'Forecast EOD' series are dropped.",
    "",
    "COVERAGE",
    f"Daily from {DATA_START.isoformat()} through yesterday - this archive continues GNI's quarterly "
    "open-data files (ireland_gas_supply_daily.xlsx / ireland_gas_demand_daily.xlsx, 2018-01-01 to "
    f"2026-03-31) rather than duplicating them. The last {RELOAD_DAYS} days are re-pulled every run.",
    "",
    "SOURCE",
    "Gas Networks Ireland data transparency: /about/data-transparency/entry-flows/physical-flows and "
    "/about/data-transparency/exit-flows/gas-consumption-by-market-sector, via the pages' own CSV export "
    "endpoint https://www.gasnetworks.ie/csv/{physicalflows,gasconsumption}?frequency=daily&date=&date_end=",
]
NOTES_SECTION_TITLES = {"UNITS", "ENTRY FLOWS", "CONSUMPTION BY SECTOR", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    stop_at = date.today() - timedelta(days=1)
    sheets = {}
    for dataset, cfg in DATASETS.items():
        existing = load_sheet(args.out, cfg["sheet"])
        start = args.start_date if existing.empty else max(args.start_date, max(existing.index) - timedelta(days=RELOAD_DAYS))
        if not existing.empty:
            # Self-heal holes inside the archive (e.g. the chunk-boundary
            # days lost before date_end's exclusivity was understood).
            have = set(existing.index)
            holes = [d.date() for d in pd.date_range(min(existing.index), max(existing.index), freq="D") if d.date() not in have]
            if holes:
                print(f"{dataset}: {len(holes)} missing day(s) inside archive, earliest {holes[0]} - refetching from there",
                      file=sys.stderr, flush=True)
                start = min(start, holes[0] - timedelta(days=1))
        print(f"{dataset}: fetching {start} to {stop_at}", file=sys.stderr, flush=True)
        new = fetch_range(dataset, start, stop_at, cfg["drop_group"])
        combined = upsert(existing, new)
        if combined.empty:
            raise RuntimeError(f"{dataset}: no data at all")
        print(f"{dataset}: {len(combined)} days, {combined.index.min()} to {combined.index.max()}", file=sys.stderr, flush=True)
        print(combined.tail(5).round(1).to_string(), file=sys.stderr, flush=True)
        sheets[cfg["sheet"]] = combined

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved to {args.out}")


if __name__ == "__main__":
    main()
