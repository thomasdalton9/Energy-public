"""
Pull the Rhine's water level at Kaub (a key navigation-depth gauge for
the whole river) from PEGELONLINE's historische-zeitreihen archive and
maintain a growing daily archive.

Data source: PEGELONLINE (German federal waterways administration),
confirmed live via RHINE_HISTORICAL_ARCHIVE_INSPECT.py:

    POST https://www.pegelonline.wsv.de/gast/historische-zeitreihen/prepare-download
    params: uuid, parameter=W, start/end (ISO-UTC), format=csv
    -> a 303 redirect (relative Location header - must be resolved
       against the host) to a ZIP containing one CSV, daily resolution.

Kaub's station UUID (1d26e504-7f9e-480a-b52c-5932be6549ab) confirmed
via RHINE_HISTORY_DEPTH_CHECK.py's station list lookup. This is a
SEPARATE endpoint from PEGELONLINE's regular /measurements.json API
(used for "current" readings elsewhere) - that one only supports a
30-day window; this historical-archive endpoint is what gives real
multi-year depth.
"""

import argparse
import io
import os
import sys
import zipfile
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

PREPARE_URL = "https://www.pegelonline.wsv.de/gast/historische-zeitreihen/prepare-download"
HOST = "https://www.pegelonline.wsv.de"
KAUB_UUID = "1d26e504-7f9e-480a-b52c-5932be6549ab"
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 60)

DATA_START = date(2000, 1, 1)  # per web search's description of this archive's depth

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "rhine_kaub_level_daily.xlsx")


def fetch_range(start, end):
    params = {
        "uuid": KAUB_UUID,
        "parameter": "W",
        "start": start.strftime("%Y-%m-%dT00:00:00.000Z"),
        "end": end.strftime("%Y-%m-%dT23:59:59.000Z"),
        "format": "csv",
    }
    r = requests.post(PREPARE_URL, headers=HEADERS, data=params, timeout=TIMEOUT, allow_redirects=False)
    r.raise_for_status()
    if r.status_code not in (301, 302, 303, 307, 308):
        raise RuntimeError(f"Expected a redirect, got HTTP {r.status_code}")
    redirect_url = r.headers["Location"]
    if redirect_url.startswith("/"):
        redirect_url = HOST + redirect_url
    r2 = requests.get(redirect_url, headers=HEADERS, timeout=TIMEOUT)
    r2.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(r2.content)) as zf:
        names = zf.namelist()
        with zf.open(names[0]) as f:
            text = f.read().decode("utf-8", errors="replace")

    df = pd.read_csv(io.StringIO(text), sep=";")
    df.columns = [c.strip() for c in df.columns]
    # Confirmed columns via a live pull: "Datum" (date) and "Wert" (value,
    # water level in cm at the gauge's local reference datum).
    df = df.rename(columns={"Datum": "date", "Wert": "level_cm"})
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["level_cm"] = pd.to_numeric(df["level_cm"], errors="coerce")
    return df.set_index("date")[["level_cm"]].sort_index()


def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new_df):
    if new_df.empty:
        return existing
    if existing.empty:
        combined = new_df
    else:
        combined = pd.concat([existing, new_df])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


NOTES_LINES = [
    "UNITS",
    "level_cm is the water level in centimeters at Kaub's gauge datum (a local reference level, not "
    "sea level) - PEGELONLINE's own published value.",
    "",
    "SCOPE",
    "Kaub is a single gauge station, not a full-river average - but it's one of the Rhine's most-",
    "watched depth points for commercial navigation (a key chokepoint for barge traffic between the "
    "Rhine's upper and lower reaches).",
    "",
    "COVERAGE",
    f"Daily since {DATA_START.isoformat()} (per PEGELONLINE's own historical archive depth).",
    "",
    "SOURCE",
    f"PEGELONLINE (German federal waterways administration)'s historische-zeitreihen archive: "
    f"{PREPARE_URL} - a different endpoint from PEGELONLINE's regular /measurements.json API (that "
    "one only supports a 30-day window).",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    existing = load_archive(args.out)
    resume_from = existing.index.max() if not existing.empty else args.start_date
    stop_at = date.today() - timedelta(days=1)

    if resume_from > stop_at:
        print(f"Archive already current ({resume_from} > {stop_at}) - nothing to do.", file=sys.stderr)
        return

    print(f"Fetching {resume_from} to {stop_at} ...", file=sys.stderr)
    new_df = fetch_range(resume_from, stop_at)
    combined = upsert(existing, new_df)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved to {args.out} ({len(combined)} days, {combined.index.min()} to {combined.index.max()})")


if __name__ == "__main__":
    main()
