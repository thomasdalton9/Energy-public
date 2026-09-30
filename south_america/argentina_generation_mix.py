"""
Pull Argentina's daily power generation mix by fuel/technology (hydro,
nuclear, thermal, biomass, solar, wind) from CAMMESA's official
"Programacion Diaria" daily report, and save daily MWh totals per
category to a CSV.

Data source: https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango
(nemo=PROGRAMACION_DIARIA), which lists one ZIP per calendar day
containing a VALORES_GENERADORES.csv - hourly output (MWh) for every
individual generating unit in the SADI grid. This script sums that
unit-level data by technology type to get a daily fuel-mix total.

Note: CAMMESA's real-time dashboards (api.cammesa.com/demanda-svc/...)
only expose a rolling current snapshot with no historical access -
even CAMMESA's own reference parser in electricitymaps-contrib
explicitly doesn't support past dates for that endpoint. This
Programacion Diaria report is the actual historical source, but there
is no bulk date-range endpoint - one ZIP is downloaded per day, so a
long --from-date range means a lot of requests and will be slow.

"Thermal" combines steam turbine, diesel, gas turbine, and combined-cycle units
(CAMMESA doesn't split thermal generation by fuel at this level, so
gas cannot be isolated from oil/diesel here). "Hydro" combines all
hydro technology types (run-of-river, reservoir, pumped storage).
"""

import argparse
import csv
import io
import sys
import time
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes/api_keys

import xlsx_notes

LOOKUP_URL = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
ATTACHMENT_URL = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
NEMO = "PROGRAMACION_DIARIA"

TIME_FMT = "%Y-%m-%dT%H:%M:%S.000Z"

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}

# Generator "type" codes (VALORES_GENERADORES.csv column 4) to category.
TYPE_TO_CATEGORY = {
    "EO": "wind",
    "BG": "biomass",
    "BM": "biomass",  # biomasa
    "TV": "thermal",  # turbina de vapor - steam turbine, not hydro
    "HR": "hydro",
    "HI": "hydro",
    "HB": "hydro",  # bombeo - pumped storage (Rio Grande)
    "MH": "hydro",  # mini-hydro
    "NU": "nuclear",
    "DI": "thermal",
    "TG": "thermal",
    "CC": "thermal",
    "FV": "solar",
    # "Importacion" (imports) is deliberately left out - not Argentine generation
}
CATEGORIES = ["hydro", "nuclear", "thermal", "biomass", "solar", "wind"]


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def find_document(session, day):
    """Find the Programacion Diaria document/attachment published for a day."""
    params = {
        "fechadesde": day.strftime(TIME_FMT),
        "fechahasta": (day + timedelta(days=1)).strftime(TIME_FMT),
        "nemo": NEMO,
    }
    response = session.get(LOOKUP_URL, params=params, timeout=30)
    response.raise_for_status()
    docs = response.json()

    expected_filename = day.strftime("PD%y%m%d.zip")
    for doc in docs:
        for attachment in doc.get("adjuntos", []):
            if attachment.get("id") == expected_filename:
                return doc, attachment
    return None, None


def download_zip(session, doc, attachment):
    params = {
        "attachmentId": attachment["id"],
        "docId": doc["id"],
        "nemo": doc.get("nemo", NEMO),
    }
    response = session.get(ATTACHMENT_URL, params=params, timeout=60)
    response.raise_for_status()
    return response.content


def parse_generators_csv(zip_bytes):
    """Sum unit-level hourly generation (MWh) by generator type code -
    kept per code (not per category) so a change to TYPE_TO_CATEGORY
    never needs a re-download."""
    totals = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        member = next(
            (name for name in zf.namelist() if name.upper().endswith("VALORES_GENERADORES.CSV")),
            None,
        )
        if member is None:
            return totals
        raw = zf.read(member).decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(raw))
        next(reader, None)  # header row
        for row in reader:
            if len(row) < 28:
                continue
            type_code = row[3].strip().strip('"')
            if not type_code:
                continue
            for value_str in row[4:28]:
                value_str = (value_str or "").strip().strip('"')
                if not value_str:
                    continue
                try:
                    totals[type_code] = totals.get(type_code, 0.0) + float(value_str)
                except ValueError:
                    continue
    return totals


def fetch_days(days):
    session = make_session()
    rows = []
    for day in days:
        print(f"Fetching {day.isoformat()}...", file=sys.stderr)
        try:
            doc, attachment = find_document(session, day)
            if doc is None:
                print(f"  no Programacion Diaria document found for {day.isoformat()}", file=sys.stderr)
            else:
                zip_bytes = download_zip(session, doc, attachment)
                totals = parse_generators_csv(zip_bytes)
                if not totals:
                    # ZIPs before 24-Oct-2024 carry only an Access .mdb, no generator CSV -
                    # keep the day as a blank row (so it isn't re-downloaded), never as zeros
                    print(f"  no generator CSV in the ZIP for {day.isoformat()}", file=sys.stderr)
                totals["date"] = day
                rows.append(totals)
        except (requests.RequestException, zipfile.BadZipFile) as exc:
            print(f"  failed for {day.isoformat()}: {exc}", file=sys.stderr)
        time.sleep(0.3)
    return rows


def to_categories(by_code):
    """Daily MWh per category from the per-type-code table."""
    unknown = sorted(set(by_code.columns) - set(TYPE_TO_CATEGORY))
    if unknown:
        totals = by_code[unknown].sum().round(0).to_dict()
        print(f"Type codes not in TYPE_TO_CATEGORY (left out of the categories): {totals}", file=sys.stderr)
    out = pd.DataFrame(index=by_code.index)
    for category in CATEGORIES:
        codes = [c for c, cat in TYPE_TO_CATEGORY.items() if cat == category and c in by_code.columns]
        out[category] = by_code[codes].sum(axis=1, min_count=1) if codes else 0.0
    # a day with no data at all stays blank rather than a day of zero generation
    out.loc[by_code.isna().all(axis=1)] = float("nan")
    return out


def load_existing(path):
    """The per-type-code table from a previous run - None if there isn't
    one (or it predates the 'By type code' tab), meaning fetch it all."""
    try:
        df = pd.read_excel(path, sheet_name=CODE_SHEET, index_col=0)
    except (FileNotFoundError, ValueError):
        return None
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    # older runs saved days without a generator CSV as all zeros - make them blank
    df.loc[df.fillna(0).eq(0).all(axis=1)] = float("nan")
    return df


NOTES_LINES = [
    "UNITS",
    "Every category column is in MWh per day - the sum of that day's 24 hourly unit-level "
    "readings (already MWh per hour), so it's a total daily energy figure, not an average MW.",
    "",
    "CATEGORIES",
    "hydro: all hydro technology types (run-of-river, reservoir, pumped storage) combined.",
    "thermal: steam turbine, diesel, gas turbine, and combined-cycle units combined - CAMMESA doesn't split "
    "thermal generation by fuel at this level, so gas cannot be isolated from oil/diesel here.",
    "nuclear, biomass, solar, wind: as published, no combining needed.",
    "",
    "SOURCE",
    "CAMMESA's official \"Programacion Diaria\" daily report - one ZIP per calendar day, "
    "summed here from unit-level generator data. CAMMESA's real-time dashboards only expose "
    "a rolling current snapshot with no historical access, which is why this script uses the "
    "report archive instead. It's the day's dispatch programme (scheduled generation), not metered output.",
    "",
    "UPDATES",
    "'By type code' keeps each day's MWh per CAMMESA generator type (TV steam turbine, TG gas turbine, CC "
    "combined cycle, DI diesel, HI/HR hydro, NU nuclear, EO wind, FV solar, BG biogas). Each run only "
    "downloads days not already in it (plus the last few, which CAMMESA can revise); 'Data' is rebuilt "
    "from it. --full re-downloads everything.",
    "History starts 24-Oct-2024: earlier ZIPs hold only an Access database, no generator CSV. Blank rows "
    "are days CAMMESA published nothing usable - they're gaps, not zero generation.",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "SOURCE", "UPDATES"}
CODE_SHEET = "By type code"
FIRST_CSV_DAY = date(2024, 10, 24)  # first Programacion Diaria ZIP with the generator CSV
REFRESH_DAYS = 3  # re-pull the latest few days each run, in case CAMMESA revises them


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from-date", type=date.fromisoformat, default=FIRST_CSV_DAY)
    parser.add_argument("--to-date", type=date.fromisoformat, default=date.today())
    parser.add_argument("--out", default="argentina_generation_mix_daily.xlsx")
    parser.add_argument("--full", action="store_true", help="ignore the saved days and download the whole range")
    args = parser.parse_args()

    existing = None if args.full else load_existing(args.out)
    wanted = [args.from_date + timedelta(days=i) for i in range((args.to_date - args.from_date).days + 1)]
    if existing is not None and len(existing):
        recent = {d for d in existing.index if d > max(existing.index) - timedelta(days=REFRESH_DAYS)}
        days = [d for d in wanted if d not in existing.index or d in recent]
        print(f"{len(existing)} days already saved; fetching {len(days)}", file=sys.stderr)
    else:
        days = wanted
    if len(days) > 60:
        print(
            f"Warning: {len(days)} days means downloading one ZIP per day from CAMMESA "
            "(there's no bulk date-range endpoint for this report) - this will be slow, once.",
            file=sys.stderr,
        )

    rows = fetch_days(days)
    new = pd.DataFrame.from_records(rows).set_index("date") if rows else pd.DataFrame()
    by_code = new if existing is None else new.combine_first(existing) if len(new) else existing
    if by_code is None or by_code.empty:
        print("No data returned.", file=sys.stderr)
        sys.exit(1)
    # a code missing on a day that has data means 0 MWh; a day with no data stays blank
    has_data = by_code.notna().any(axis=1)
    by_code = by_code.sort_index()
    by_code = by_code[by_code.index >= min(wanted)]
    has_data = has_data.reindex(by_code.index)
    by_code.loc[has_data] = by_code.loc[has_data].fillna(0.0)
    by_code.index.name = "date"
    by_code = by_code[sorted(by_code.columns)]
    df = to_categories(by_code)

    xlsx_notes.write_workbook(args.out, {"Data": df, CODE_SHEET: by_code}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(df)} days to {args.out} (Data, {CODE_SHEET} + Units tabs)")
    print(df.tail())


if __name__ == "__main__":
    main()
