"""
Pull Argentina's daily power generation mix by fuel/technology (hydro,
nuclear, thermal, biomass, solar, wind) from CAMMESA's official
"Programacion Diaria" daily report, plus - from the same daily ZIP,
previously downloaded but not parsed - a regional breakdown of
generation, demand, and thermal fuel consumption.

Data source: https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango
(nemo=PROGRAMACION_DIARIA), which lists one ZIP per calendar day
containing several CSVs:
  - VALORES_GENERADORES.csv - hourly output (MWh) for every individual
    generating unit in the SADI grid, WITH a REGION column (its first
    column, previously unread) - Argentina's 10 wholesale-market
    regions: BAS (Buenos Aires), CEN (Centro), COM (Comahue), CTMSG
    (Salto Grande, the Argentina/Uruguay binational plant), CUY (Cuyo),
    GBA (Gran Buenos Aires), LIT (Litoral), NEA (Noreste), NOA
    (Noroeste), PAT (Patagonia).
  - BALANCE.csv - each region's generation-side (TIPO='G': Nuclear,
    Termica, 'Ren Hidro >50MW', 'Ren ley 26190' - renewables under Law
    26190, i.e. wind/solar/biomass/small hydro bundled) and demand-side
    (TIPO='D': Demanda Neta, Perdidas, Bombeo, Importacion) totals - a
    coarser, independent regional cross-check of the same day.
  - CONSUMOS_COMBUSTIBLES.csv - thermal fuel consumption by region,
    plant, and fuel-type code - not published every day (only when
    CAMMESA has revised/reported it), and its physical unit isn't
    disclosed in the file itself, so raw values are kept as-is rather
    than relabeled into something that might be wrong.

This script sums unit-level/regional data by day to get daily totals.
CAMMESA's real-time dashboards (api.cammesa.com/demanda-svc/...) only
expose a rolling current snapshot with no historical access - even
CAMMESA's own reference parser in electricitymaps-contrib explicitly
doesn't support past dates for that endpoint. This Programacion Diaria
report is the actual historical source, but there is no bulk date-range
endpoint - one ZIP is downloaded per day, so a long --from-date range
means a lot of requests and will be slow.

"Thermal" combines steam turbine, diesel, gas turbine, and combined-cycle units
(CAMMESA doesn't split thermal generation by fuel at this level, so
gas cannot be isolated from oil/diesel here in the main category - see
"Fuel consumption by region" for a fuel-code-level breakdown instead).
"Hydro" combines all hydro technology types (run-of-river, reservoir,
pumped storage).
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
COUNTRY = "Argentina"

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

# Thermal fuel-type codes (CONSUMOS_COMBUSTIBLES.csv column 3) - best-effort
# labels; a code not in this dict keeps its raw code as the label (printed
# as a warning) rather than being dropped, since this is supplementary
# detail, not the main generation categorization.
FUEL_CODE_LABELS = {
    "GN": "Gas Natural",
    "GO": "Gasoil/Diesel",
    "FO": "Fuel Oil",
    "CM": "Carbon Mineral",
    "U2": "Uranio (combustible nuclear)",
    "UE": "Uranio enriquecido (combustible nuclear)",
    "BC": "Biocombustible",
}


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


def _sum_hours(totals, key, hour_values):
    for value_str in hour_values:
        value_str = (value_str or "").strip().strip('"')
        if not value_str:
            continue
        try:
            totals[key] = totals.get(key, 0.0) + float(value_str)
        except ValueError:
            continue


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
            _sum_hours(totals, type_code, row[4:28])
    return totals


def parse_generators_by_region(zip_bytes):
    """Same file as parse_generators_csv, but keyed by (region, type
    code) using column 0 (REGION), previously unread."""
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
        next(reader, None)
        for row in reader:
            if len(row) < 28:
                continue
            region = row[0].strip().strip('"')
            type_code = row[3].strip().strip('"')
            if not region or not type_code:
                continue
            _sum_hours(totals, f"{region}|{type_code}", row[4:28])
    return totals


def parse_balance_csv(zip_bytes):
    """Regional generation/demand balance from BALANCE.csv (TIPO, RGE,
    VARIABLE, H01..H24) - keyed by 'REGION|VARIABLE'. Not every day's
    ZIP has this file (older/edge-case days may omit it)."""
    totals = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        member = next((n for n in zf.namelist() if n.upper() == "BALANCE.CSV"), None)
        if member is None:
            return totals
        raw = zf.read(member).decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(raw))
        next(reader, None)
        for row in reader:
            if len(row) < 27:
                continue
            rge = row[1].strip().strip('"')
            variable = row[2].strip().strip('"')
            if not rge or not variable:
                continue
            _sum_hours(totals, f"{rge}|{variable}", row[3:27])
    return totals


def parse_combustibles_csv(zip_bytes):
    """Thermal fuel consumption by region, plant, and fuel-type code
    from CONSUMOS_COMBUSTIBLES.csv (REGION, CENTRAL, TIPO_COMB,
    H01..H24) - keyed by 'REGION|CENTRAL|FUEL_CODE'. Only present in
    some days' ZIPs (CAMMESA doesn't publish it every day)."""
    totals = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        member = next((n for n in zf.namelist() if n.upper() == "CONSUMOS_COMBUSTIBLES.CSV"), None)
        if member is None:
            return totals
        raw = zf.read(member).decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(raw))
        next(reader, None)
        for row in reader:
            if len(row) < 27:
                continue
            region = row[0].strip().strip('"')
            central = row[1].strip().strip('"')
            fuel_code = row[2].strip().strip('"')
            if not region or not central or not fuel_code:
                continue
            _sum_hours(totals, f"{region}|{central}|{fuel_code}", row[3:27])
    return totals


def fetch_days(days):
    session = make_session()
    by_code_rows, by_region_rows, balance_rows, fuel_rows = [], [], [], []
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
                by_code_rows.append(totals)

                region_totals = parse_generators_by_region(zip_bytes)
                region_totals["date"] = day
                by_region_rows.append(region_totals)

                balance_totals = parse_balance_csv(zip_bytes)
                balance_totals["date"] = day
                balance_rows.append(balance_totals)

                fuel_totals = parse_combustibles_csv(zip_bytes)
                fuel_totals["date"] = day
                fuel_rows.append(fuel_totals)
        except (requests.RequestException, zipfile.BadZipFile) as exc:
            print(f"  failed for {day.isoformat()}: {exc}", file=sys.stderr)
        time.sleep(0.3)
    return by_code_rows, by_region_rows, balance_rows, fuel_rows


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


def load_existing_wide(path, sheet_name):
    """A previously-saved wide table (index=date, columns=raw keys) -
    None if there isn't one, meaning fetch it all."""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, index_col=0)
    except (FileNotFoundError, ValueError):
        return None
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def combine_wide(rows, existing, wanted_from):
    """Build/merge a wide table (index=date) from freshly-fetched rows
    and whatever was already saved, the same combine-and-refresh logic
    for all four CAMMESA-derived tables in this script."""
    new = pd.DataFrame.from_records(rows).set_index("date") if rows else pd.DataFrame()
    wide = new if existing is None else new.combine_first(existing) if len(new) else existing
    if wide is None or wide.empty:
        return wide
    has_data = wide.notna().any(axis=1)
    wide = wide.sort_index()
    wide = wide[wide.index >= wanted_from]
    has_data = has_data.reindex(wide.index)
    wide.loc[has_data] = wide.loc[has_data].fillna(0.0)
    wide.index.name = "date"
    return wide[sorted(wide.columns)]


def melt_region_generation(by_region):
    """'REGION|TYPE_CODE' wide columns -> long format (date, region,
    category, country, mwh), rolling type codes into the same
    CATEGORIES used in the national 'Data' tab."""
    if by_region is None or by_region.empty:
        return pd.DataFrame(columns=["date", "region", "category", "country", "mwh"])
    records = []
    for col in by_region.columns:
        region, type_code = col.split("|", 1)
        category = TYPE_TO_CATEGORY.get(type_code)
        if category is None:
            continue  # e.g. "Importacion" - not Argentine generation, same as the national tab
        for dt, value in by_region[col].items():
            if pd.notna(value):
                records.append({"date": dt, "region": region, "category": category, "country": COUNTRY, "mwh": value})
    if not records:
        return pd.DataFrame(columns=["date", "region", "category", "country", "mwh"])
    out = pd.DataFrame(records)
    return out.groupby(["date", "region", "category", "country"], as_index=False)["mwh"].sum()


def melt_balance(by_balance):
    """'REGION|VARIABLE' wide columns -> long format (date, region,
    side, variable, country, mwh). side is 'generation' for the four
    generation-side VARIABLE values, 'demand' for the other four."""
    if by_balance is None or by_balance.empty:
        return pd.DataFrame(columns=["date", "region", "side", "variable", "country", "mwh"])
    generation_variables = {"Nuclear", "Termica", "Ren Hidro >50MW", "Ren ley 26190"}
    records = []
    for col in by_balance.columns:
        region, variable = col.split("|", 1)
        side = "generation" if variable in generation_variables else "demand"
        for dt, value in by_balance[col].items():
            if pd.notna(value):
                records.append({"date": dt, "region": region, "side": side, "variable": variable,
                                 "country": COUNTRY, "mwh": value})
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["date", "region", "side", "variable", "country", "mwh"])


def melt_fuel_consumption(by_fuel):
    """'REGION|CENTRAL|FUEL_CODE' wide columns -> long format (date,
    region, plant, fuel_code, fuel_label, country, value). Units as
    published by CAMMESA - not disclosed in the source file, kept raw."""
    if by_fuel is None or by_fuel.empty:
        return pd.DataFrame(columns=["date", "region", "plant", "fuel_code", "fuel_label", "country", "value"])
    unknown_codes = set()
    records = []
    for col in by_fuel.columns:
        region, plant, fuel_code = col.split("|", 2)
        label = FUEL_CODE_LABELS.get(fuel_code)
        if label is None:
            unknown_codes.add(fuel_code)
            label = fuel_code
        for dt, value in by_fuel[col].items():
            if pd.notna(value):
                records.append({"date": dt, "region": region, "plant": plant, "fuel_code": fuel_code,
                                 "fuel_label": label, "country": COUNTRY, "value": value})
    if unknown_codes:
        print(f"Fuel codes not in FUEL_CODE_LABELS (kept, raw code used as label): {sorted(unknown_codes)}",
              file=sys.stderr)
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["date", "region", "plant", "fuel_code", "fuel_label", "country", "value"])


NOTES_LINES = [
    "UNITS",
    "'Data'/'By type code'/'By region': MWh per day - the sum of that day's 24 hourly unit-level "
    "readings (already MWh per hour), so it's a total daily energy figure, not an average MW. "
    "'Regional balance' uses the same sum-of-24-hourly-readings convention. 'Fuel consumption by "
    "region': raw values as published by CAMMESA - the physical unit isn't disclosed in the source "
    "file, so nothing is assumed or converted here.",
    "",
    "CATEGORIES",
    "hydro: all hydro technology types (run-of-river, reservoir, pumped storage) combined.",
    "thermal: steam turbine, diesel, gas turbine, and combined-cycle units combined - CAMMESA doesn't split "
    "thermal generation by fuel at this level in the main categorization (see 'Fuel consumption by region' "
    "for a fuel-code-level breakdown instead).",
    "nuclear, biomass, solar, wind: as published, no combining needed.",
    "",
    "REGIONS",
    "Argentina's 10 CAMMESA wholesale-market regions: BAS (Buenos Aires), CEN (Centro), COM (Comahue), "
    "CTMSG (Salto Grande, the Argentina/Uruguay binational plant), CUY (Cuyo), GBA (Gran Buenos Aires), "
    "LIT (Litoral), NEA (Noreste), NOA (Noroeste), PAT (Patagonia). 'By region' uses the same REGION field "
    "and TYPE_TO_CATEGORY mapping as the national 'Data'/'By type code' tabs, just also split by region. "
    "'Regional balance' is CAMMESA's own coarser, independent regional generation/demand split (from "
    "BALANCE.csv) - a cross-check, not identical categories to 'By region'. 'Fuel consumption by region' is "
    "not published every day (only when CAMMESA has revised/reported it) - a region with no rows on a given "
    "day means no data was published that day, not zero consumption.",
    "",
    "SOURCE",
    "CAMMESA's official \"Programacion Diaria\" daily report - one ZIP per calendar day, "
    "summed here from unit-level generator data plus BALANCE.csv and CONSUMOS_COMBUSTIBLES.csv from "
    "the same ZIP. CAMMESA's real-time dashboards only expose a rolling current snapshot with no "
    "historical access, which is why this script uses the report archive instead. It's the day's "
    "dispatch programme (scheduled generation), not metered output.",
    "",
    "UPDATES",
    "'By type code' keeps each day's MWh per CAMMESA generator type (TV steam turbine, TG gas turbine, CC "
    "combined cycle, DI diesel, HI/HR hydro, NU nuclear, EO wind, FV solar, BG biogas). Each run only "
    "downloads days not already in it (plus the last few, which CAMMESA can revise); 'Data'/'By region'/"
    "'Regional balance'/'Fuel consumption by region' are rebuilt from their own cached wide tables the same "
    "way. --full re-downloads everything.",
    "History starts 24-Oct-2024: earlier ZIPs hold only an Access database, no generator CSV. Blank rows "
    "are days CAMMESA published nothing usable - they're gaps, not zero generation.",
]
NOTES_SECTION_TITLES = {"UNITS", "CATEGORIES", "REGIONS", "SOURCE", "UPDATES"}
CODE_SHEET = "By type code"
REGION_CODE_SHEET = "By region (raw)"
BALANCE_CODE_SHEET = "Regional balance (raw)"
FUEL_CODE_SHEET = "Fuel consumption (raw)"
FIRST_CSV_DAY = date(2024, 10, 24)  # first Programacion Diaria ZIP with the generator CSV
REFRESH_DAYS = 3  # re-pull the latest few days each run, in case CAMMESA revises them


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from-date", type=date.fromisoformat, default=FIRST_CSV_DAY)
    parser.add_argument("--to-date", type=date.fromisoformat, default=date.today())
    parser.add_argument("--out", default="argentina_generation_mix_daily.xlsx")
    parser.add_argument("--full", action="store_true", help="ignore the saved days and download the whole range")
    args = parser.parse_args()

    existing_code = None if args.full else load_existing_wide(args.out, CODE_SHEET)
    existing_region = None if args.full else load_existing_wide(args.out, REGION_CODE_SHEET)
    existing_balance = None if args.full else load_existing_wide(args.out, BALANCE_CODE_SHEET)
    existing_fuel = None if args.full else load_existing_wide(args.out, FUEL_CODE_SHEET)
    # older runs saved days without a generator CSV as all zeros - make them blank
    if existing_code is not None and len(existing_code):
        existing_code.loc[existing_code.fillna(0).eq(0).all(axis=1)] = float("nan")

    wanted = [args.from_date + timedelta(days=i) for i in range((args.to_date - args.from_date).days + 1)]
    if existing_code is not None and len(existing_code):
        recent = {d for d in existing_code.index if d > max(existing_code.index) - timedelta(days=REFRESH_DAYS)}
        days = [d for d in wanted if d not in existing_code.index or d in recent]
        print(f"{len(existing_code)} days already saved; fetching {len(days)}", file=sys.stderr)
    else:
        days = wanted
    if len(days) > 60:
        print(
            f"Warning: {len(days)} days means downloading one ZIP per day from CAMMESA "
            "(there's no bulk date-range endpoint for this report) - this will be slow, once.",
            file=sys.stderr,
        )

    code_rows, region_rows, balance_rows, fuel_rows = fetch_days(days)
    by_code = combine_wide(code_rows, existing_code, min(wanted))
    if by_code is None or by_code.empty:
        print("No data returned.", file=sys.stderr)
        sys.exit(1)
    by_region = combine_wide(region_rows, existing_region, min(wanted))
    by_balance = combine_wide(balance_rows, existing_balance, min(wanted))
    by_fuel = combine_wide(fuel_rows, existing_fuel, min(wanted))

    df = to_categories(by_code)
    region_long = melt_region_generation(by_region)
    balance_long = melt_balance(by_balance)
    fuel_long = melt_fuel_consumption(by_fuel)

    sheets = {
        "Data": df,
        CODE_SHEET: by_code,
        "By region": region_long,
        REGION_CODE_SHEET: by_region if by_region is not None else pd.DataFrame(),
        "Regional balance": balance_long,
        BALANCE_CODE_SHEET: by_balance if by_balance is not None else pd.DataFrame(),
        "Fuel consumption by region": fuel_long,
        FUEL_CODE_SHEET: by_fuel if by_fuel is not None else pd.DataFrame(),
    }
    xlsx_notes.write_workbook(args.out, sheets, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(df)} days to {args.out} "
          f"(Data, By region, Regional balance, Fuel consumption by region + raw/Units tabs)")
    print(df.tail())
    print(f"\nBy region sample:\n{region_long.tail()}")


if __name__ == "__main__":
    main()
