"""
Trinidad & Tobago natural gas production (by company) and utilization
(by sector) from the Ministry of Energy and Energy Industries (MEEI)'s
Consolidated Monthly Bulletin - free, public, no key.

REWRITTEN to fix the original version's "needs manual URL updates"
limitation and replace brittle PDF table extraction:
  - The bulletin LISTING page (energy.gov.tt/?p=11949) always links the
    CURRENT bulletin directly - confirmed live via TRINIDAD_GAS_EXCEL_
    DISCOVERY.py - so the current bulletin's URL is resolved from that
    page on every run instead of being hardcoded.
  - Bulletins also publish in Excel (.xlsx) alongside the PDF - far
    more reliable to parse than PDF table extraction (which needed a
    real workaround for a pdfplumber table-merge quirk in the original
    version). Confirmed via TRINIDAD_GAS_XLSX_INSPECT.py: sheet
    "3A,3B" has both tables with clean cell values, no OCR/extraction
    ambiguity, just occasional stray data-quality issues in the sheet
    itself (see the TOTAL row note below).

TABLE 3A - Natural Gas Production by Company (MMSCF/D)
TABLE 3B - Natural Gas Utilization by Sector (MMSCF/D)
Both are monthly, year-to-date within the bulletin (e.g. a bulletin
covering January-May has non-zero columns for Jan-May only, 0 for the
rest of that year).

NOTE: the sheet's own "AVG <year>" column for the Utilization TOTAL row
was found to be wrong in a live pull (947.58, when the monthly values
themselves average to ~2274) - a data-quality issue in MEEI's own
spreadsheet, not a parsing bug. This script reads only the real monthly
columns and does not use the sheet's own AVG column at all, to avoid
propagating that kind of error.

Outputs (trinidad_gas.xlsx):
  Production by company    date, company, mmscfd, country
  Utilization by sector    date, sector, mmscfd, country

Usage: python3 TRINIDAD_GAS.py [--out trinidad_gas.xlsx] [--url <bulletin .xlsx URL>]
"""
print("STARTING", flush=True)

import argparse
import io
import re
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

LISTING_URL = "https://www.energy.gov.tt/?p=11949"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 120)
COUNTRY = "Trinidad and Tobago"
OUT_DEFAULT = "trinidad_gas.xlsx"
GAS_SHEET_CANDIDATES = ["3A,3B", "3A, 3B", "3A 3B"]


def resolve_current_bulletin_url():
    """The listing page always links the current bulletin directly (in
    both PDF and Excel) - confirmed live via TRINIDAD_GAS_EXCEL_
    DISCOVERY.py. Picks the newest-dated .xlsx link found (bulletin
    filenames embed the publish date at the end, e.g.
    "...-16-06-2026.xlsx")."""
    r = requests.get(LISTING_URL, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    xlsx_links = sorted(set(re.findall(r'href="([^"]+\.xlsx?)"', r.text, re.I)))
    bulletin_links = [link for link in xlsx_links if "Consolidated-Monthly-Bulletin" in link]
    if not bulletin_links:
        raise RuntimeError(f"No Excel bulletin link found on {LISTING_URL} - page layout may have changed")

    def publish_date_key(url):
        m = re.search(r"(\d{2})-(\d{2})-(\d{4})\.xlsx?$", url)
        return (m.group(3), m.group(2), m.group(1)) if m else ("0000", "00", "00")

    return max(bulletin_links, key=publish_date_key)


def find_gas_sheet(xl):
    for name in xl.sheet_names:
        if name.strip() in GAS_SHEET_CANDIDATES:
            return name
    # Fall back to scanning every sheet's first column for the table title,
    # in case MEEI ever renames the tab.
    for name in xl.sheet_names:
        df = xl.parse(name, header=None, nrows=5)
        if df.iloc[:, 0].astype(str).str.contains("PRODUCTION BY COMPANY", case=False, na=False).any():
            return name
    raise RuntimeError(f"Could not find the gas production/utilization sheet among: {xl.sheet_names}")


def find_tables(df):
    """df is the raw sheet (header=None). Returns (production_rows,
    utilization_rows) - each a sub-DataFrame starting at its title row
    up to (not including) the next title row or a blank row."""
    col0 = df.iloc[:, 0].astype(str)
    production_idx = col0[col0.str.contains("PRODUCTION BY COMPANY", case=False, na=False)].index
    utilization_idx = col0[col0.str.contains("UTILI[SZ]ATION BY SECTOR", case=False, na=False, regex=True)].index
    if len(production_idx) == 0 or len(utilization_idx) == 0:
        raise RuntimeError("Could not locate both table titles in the gas sheet")
    p_start, u_start = production_idx[0], utilization_idx[0]
    return df.iloc[p_start:u_start], df.iloc[u_start:]


def table_to_long(table, entity_col):
    """table row 0 = title, row 1 = header (entity name, then one
    datetime column per month, then an AVG column we deliberately
    ignore - see module docstring), row 2+ = data, ending at TOTAL."""
    header = table.iloc[1]
    month_cols = [i for i, val in enumerate(header) if isinstance(val, pd.Timestamp)]
    rows = []
    for _, row in table.iloc[2:].iterrows():
        entity = row.iloc[0]
        if not isinstance(entity, str) or not entity.strip():
            continue
        entity = entity.strip()
        if entity.upper().startswith("NOTE"):
            break
        for i in month_cols:
            value = row.iloc[i]
            if pd.isna(value):
                continue
            value = float(value)
            if value == 0:
                continue  # zero = month not yet reached this year, not a real reading
            date = pd.Timestamp(header.iloc[i]).replace(day=1)
            rows.append({"date": date, entity_col: entity, "mmscfd": value, "country": COUNTRY})
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--url", default=None, help="override the bulletin .xlsx URL (skips auto-discovery)")
    args = parser.parse_args()

    bulletin_url = args.url or resolve_current_bulletin_url()
    print(f"Downloading MEEI bulletin: {bulletin_url}", flush=True)
    r = requests.get(bulletin_url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()

    xl = pd.ExcelFile(io.BytesIO(r.content))
    sheet_name = find_gas_sheet(xl)
    print(f"Using sheet: {sheet_name!r}", flush=True)
    full = xl.parse(sheet_name, header=None)
    production_table, utilization_table = find_tables(full)

    production = table_to_long(production_table, "company")
    utilization = table_to_long(utilization_table, "sector")

    print(f"Production: {len(production)} rows, {production['date'].nunique() if not production.empty else 0} months",
          flush=True)
    print(f"Utilization: {len(utilization)} rows, "
          f"{utilization['date'].nunique() if not utilization.empty else 0} months", flush=True)
    if not utilization.empty:
        print(utilization[utilization['sector'] != 'TOTAL'].pivot(index='date', columns='sector', values='mmscfd')
              .to_string(), flush=True)

    notes = [
        "UNITS",
        "MMSCF/D - million standard cubic feet per day, as published (a daily-average rate for each month, not "
        "a monthly total volume).",
        "",
        "SECTORS",
        "'Utilization by sector' breaks down where Trinidad's gas actually goes - Power Generation is just one "
        "of eleven categories, alongside Ammonia Manufacture, Methanol Manufacture, Refinery, Iron & Steel, "
        "Cement, Ammonia Derivatives (Urea/UAN/Melamine), Gas Processing, Small Consumers, and LNG (by far the "
        "largest single use). 'TOTAL' is MEEI's own published system total, kept as its own row.",
        "",
        "COVERAGE",
        "The bulletin's own listing page (energy.gov.tt) always links the current bulletin directly, so this "
        "resolves the current bulletin's URL fresh on every run - no manual updates needed. Each bulletin "
        "covers the current year to date; months not yet reached that year show as 0 in the source and are "
        "dropped here rather than kept as fake zero readings.",
        "",
        "DATA QUALITY NOTE",
        "The source spreadsheet's own 'AVG <year>' column for the Utilization TOTAL row does not match the "
        "average of that row's monthly values in at least one bulletin seen live - this script does not read "
        "that column at all (only the real monthly columns), to avoid propagating that kind of error.",
        "",
        "SOURCE",
        f"Ministry of Energy and Energy Industries (MEEI) Consolidated Monthly Bulletin: {LISTING_URL}",
    ]
    sheets = {"Production by company": production, "Utilization by sector": utilization}
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "COVERAGE", "DATA QUALITY NOTE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
