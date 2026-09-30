"""
Trinidad & Tobago natural gas production (by company) and utilization
(by sector) from the Ministry of Energy and Energy Industries (MEEI)'s
Consolidated Monthly Bulletin - free, public, no key.

Confirmed via TRINIDAD_GAS_PDF_INSPECT.py: the bulletin PDF has two
clean, machine-parseable tables on one page:
  TABLE 3A - Natural Gas Production by Company (MMSCF/D)
  TABLE 3B - Natural Gas Utilization by Sector (MMSCF/D)
Both are monthly, year-to-date within the bulletin (e.g. a bulletin
titled "January-April 2025" has non-zero columns for Jan-Apr only).

data.gov.tt (the structured CSV alternative) was down for maintenance
when checked - this PDF is the only confirmed-live source right now.

IMPORTANT LIMITATION: only ONE bulletin URL is confirmed (the most
recent one found via web search). MEEI publishes these periodically
under a date-stamped filename - there's no confirmed index/archive page
listing every past bulletin, so this script can't backfill history on
its own. BULLETIN_URL needs updating by hand to a newer bulletin as MEEI
publishes them (each new one still contains the full current year to
date, so re-running against a newer URL naturally extends coverage).

Outputs (trinidad_gas.xlsx):
  Production by company    date, company, mmscfd, country
  Utilization by sector    date, sector, mmscfd, country

Usage: python3 TRINIDAD_GAS.py [--out trinidad_gas.xlsx]
"""
print("STARTING", flush=True)

import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

BULLETIN_URL = ("https://www.energy.gov.tt/wp-content/uploads/2025/09/"
                 "MEEI-Consolidated-Monthly-Bulletins_January-April-2025-14-07-2025.pdf")
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 120)
COUNTRY = "Trinidad and Tobago"
OUT_DEFAULT = "trinidad_gas.xlsx"

MONTH_COL_RE = re.compile(r"^[A-Za-z]{3}-\d{2}$")  # e.g. "Jan-25"


def to_number(val):
    if val in (None, "", "0"):
        return None
    try:
        return float(str(val).replace(",", ""))
    except ValueError:
        return None


def find_tables(pdf):
    """Returns (production_table, utilization_table) - the raw row lists
    from pdfplumber, located by title rather than a fixed page number
    (bulletins can reorder pages between issues).

    pdfplumber merges Table 3A and 3B into a SINGLE extracted table when
    there's no visible border between them on the page - confirmed live
    (a first version of this function assumed one extracted table per
    logical table and silently missed 3B entirely). So this scans EVERY
    extracted table's rows for a title cell, and splits on it wherever
    one appears mid-table."""
    production, utilization = None, None
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table:
                continue
            split_at = None
            for i, row in enumerate(table):
                title = str((row[0] if row else "") or "").upper()
                if "PRODUCTION BY COMPANY" in title:
                    production = table[i:] if split_at is None else production
                    split_at = i
                elif "UTILIZATION BY SECTOR" in title or "UTILISATION BY SECTOR" in title:
                    if production is not None and split_at is not None:
                        production = table[split_at:i]
                    utilization = table[i:]
    return production, utilization


def table_to_long(table, entity_col):
    """table[0] = title row, table[1] = header (entity name, Jan-25, ..., AVG YYYY),
    table[2:] = data rows, last row = TOTAL (kept as its own entity)."""
    header = table[1]
    month_cols = [(i, col) for i, col in enumerate(header) if col and MONTH_COL_RE.match(col.strip())]
    rows = []
    for row in table[2:]:
        if not row or not row[0]:
            continue
        entity = row[0].strip()
        for i, month_label in month_cols:
            value = to_number(row[i]) if i < len(row) else None
            if value is not None:
                date = pd.to_datetime(month_label, format="%b-%y")
                rows.append({"date": date, entity_col: entity, "mmscfd": value, "country": COUNTRY})
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--url", default=BULLETIN_URL, help="override the bulletin PDF URL")
    args = parser.parse_args()

    print(f"Downloading MEEI bulletin: {args.url}", flush=True)
    r = requests.get(args.url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()

    import pdfplumber
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        production_table, utilization_table = find_tables(pdf)

    if production_table is None:
        print("  WARNING: 'Production by Company' table not found - page layout may have changed", flush=True)
    if utilization_table is None:
        print("  WARNING: 'Utilization by Sector' table not found - page layout may have changed", flush=True)

    production = table_to_long(production_table, "company") if production_table else pd.DataFrame()
    utilization = table_to_long(utilization_table, "sector") if utilization_table else pd.DataFrame()

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
        "LIMITATION",
        "Only one bulletin is confirmed reachable right now (data.gov.tt, the structured CSV alternative, is "
        "down for maintenance) - this script has no way to discover or backfill older bulletins on its own. "
        "BULLETIN_URL needs updating by hand as MEEI publishes newer ones; each new bulletin still covers the "
        "full current year to date, so history grows a little each time this is pointed at a newer URL.",
        "",
        "SOURCE",
        "Ministry of Energy and Energy Industries (MEEI) Consolidated Monthly Bulletin, energy.gov.tt.",
    ]
    sheets = {"Production by company": production, "Utilization by sector": utilization}
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "LIMITATION", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
