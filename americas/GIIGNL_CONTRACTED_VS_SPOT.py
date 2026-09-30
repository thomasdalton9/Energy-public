"""
Pull GIIGNL's annual LNG Industry report PDFs and extract, per exporting
country/basin and per report year, contracted (long-term) vs spot/
short-term export volumes.

Data source: GIIGNL (Groupe International des Importateurs de Gaz
Naturel Liquefie) annual reports, published as public PDFs at
https://giignl-documents.s3.fr-par.scw.cloud/public/ar-{year}-annual-report.pdf
Found via GIIGNL_DISCOVERY.py / GIIGNL_PDF_INSPECT.py - no API, no CSV/
Excel, PDF only, but with genuinely real (non-scanned) extractable text
and two data tables that together answer contracted-vs-spot:

  1. Total exports by exporting country/basin (a normal, cleanly-
     extractable table via pdfplumber's table detection - "ATLANTIC
     BASIN1 / MIDDLE EAST1 / PACIFIC BASIN1" sections, each row
     Country, MT, Global Share, Var Mt, Var %).
  2. A "Spot and Short-Term* LNG Quantities (in MT) received in {year}"
     bilateral matrix (importing market x exporting country/basin) -
     footnoted "*Quantities delivered under contracts of a duration of
     4 years or less". Its own GRAND TOTAL row, read per exporting
     country/basin column, is that country's total spot+short-term
     EXPORT volume for the year.

Contracted (long-term) = Total exports (table 1) - Spot/short-term
exports (table 2's grand total row), per exporting country/basin.

PARSING APPROACH: table 1 extracts cleanly via pdfplumber's normal
extract_tables(). Table 2 does not - pdfplumber's column-boundary
detection merges some countries' values into single jammed strings
(confirmed by actually inspecting the raw extraction), and a country
with an exact-zero total for some column renders as a blank rather
than "0.0", breaking any approach that just zips numbers to headers by
position/count. Both tables are instead read via extract_words() (word-
level, with x-coordinates) and matched by NEAREST X-COORDINATE between
each header word's column and each grand-total-row number's column -
robust to a missing/blank cell (that header's value is just left as
NaN) and to neighbouring columns' text visually overlapping.

Page numbers are NOT hardcoded (confirmed to shift across report
editions) - each year's PDF is scanned for the two tables by their own
title text instead.

Only as many years as are actually found at the known URL pattern are
pulled - not assumed to run back further than GIIGNL_DISCOVERY.py found
real links for (2020 onward, as of this script's writing).
"""

import argparse
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

BASE_URL = "https://giignl-documents.s3.fr-par.scw.cloud/public/ar-{year}-annual-report.pdf"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [10, 30]

# Confirmed present via GIIGNL_DISCOVERY.py; a year missing at the URL
# (404) is skipped rather than treated as an error - GIIGNL's own first
# published edition is the natural start of this history, not a fixed
# assumption made here.
CANDIDATE_YEARS = range(2020, 2031)

# Country/basin names as they appear in GIIGNL's own export tables -
# used only to recognise which table-1 rows are basin subtotals
# (skipped) vs real countries (kept), not to filter which countries are
# tracked - whatever GIIGNL reports shows up as a row automatically.
BASIN_SUBTOTAL_LABELS = {"ATLANTIC BASIN1", "MIDDLE EAST1", "PACIFIC BASIN1",
                          "ATLANTIC BASIN", "MIDDLE EAST", "PACIFIC BASIN"}

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "giignl_contracted_vs_spot_annual.xlsx")


def fetch_pdf_bytes(year):
    url = BASE_URL.format(year=year)
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS[attempt - 1])
    raise last_error


def find_export_totals_table(pdf):
    """Table 1: total exports by exporting country/basin. Scans every
    page's cleanly-extracted tables for one whose first column contains
    a basin subtotal label (ATLANTIC BASIN.../MIDDLE EAST.../PACIFIC
    BASIN...) - the reliable fingerprint of this table, since its exact
    page number moves between report editions."""
    rows = {}
    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                if not row or not row[0]:
                    continue
                label = re.sub(r"\s+", " ", str(row[0])).strip().upper()
                is_subtotal = any(label.startswith(b) for b in BASIN_SUBTOTAL_LABELS)
                # A real country/basin row in this table has a numeric MT value
                # in the second column - filters out unrelated tables that also
                # happen to start with a country-like first cell.
                if len(row) < 2 or row[1] is None:
                    continue
                try:
                    mt = float(str(row[1]).replace(",", ""))
                except ValueError:
                    continue
                if is_subtotal or label in rows or _looks_like_export_country_row(row):
                    rows[row[0].strip()] = mt
    return rows


def _looks_like_export_country_row(row):
    """A real data row has MT, a %, a +/- variance Mt and a +/- variance
    % - four numeric-ish cells after the label, matching table 1's
    known column layout. Cheap enough to just require >= 3 non-empty
    cells after the label instead of parsing every one."""
    return sum(1 for c in row[1:] if c not in (None, "")) >= 3


def find_spot_shortterm_grand_total(pdf):
    """Table 2: the bilateral spot/short-term matrix's own GRAND TOTAL
    row, read by matching each number's x-coordinate to the nearest
    header word's x-coordinate - robust to pdfplumber's column-merging
    on the data rows (not relied on here) and to a column whose total
    happens to be exactly zero (rendered blank, not "0.0" - left as NaN
    instead of silently shifting every later column)."""
    for page in pdf.pages:
        text = page.extract_text() or ""
        if "Spot and Short" not in text and "SPOT AND SHORT" not in text.upper():
            continue
        if "GRAND TOTAL" not in text.upper().replace("\n", " "):
            # Some editions may wrap the title differently - still require
            # an actual grand total row to be present before trusting this page.
            continue

        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        lines = _group_words_into_lines(words)

        header_line = _find_header_line(lines)
        total_line = _find_grand_total_line(lines)
        if header_line is None or total_line is None:
            continue

        headers = _merge_wrapped_header_words(header_line)
        totals = [w for w in total_line if _is_number(w["text"])]

        result = {}
        for h in headers:
            nearest = min(totals, key=lambda w: abs(_x_center(w) - _x_center(h)), default=None)
            if nearest is not None:
                result[h["text"]] = float(nearest["text"].replace(",", ""))
        return result
    return None


def _x_center(word):
    return (word["x0"] + word["x1"]) / 2


def _is_number(s):
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", s.replace(",", "")))


def _group_words_into_lines(words, y_tolerance=3):
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        placed = False
        for line in lines:
            if abs(line[0]["top"] - w["top"]) <= y_tolerance:
                line.append(w)
                placed = True
                break
        if not placed:
            lines.append([w])
    return lines


def _find_header_line(lines):
    for line in lines:
        texts = [w["text"] for w in line]
        if "Markets" in texts:
            return sorted(line, key=lambda w: w["x0"])
    return None


def _find_grand_total_line(lines):
    joined = [(l, "".join(w["text"] for w in l).upper()) for l in lines]
    for line, joined_text in joined:
        # GIIGNL's own PDF renders this row's label with doubled letters
        # (a font/kerning artifact seen in every inspected page, not a
        # typo here) - "GGRRAANNDD TTOOTTAALL" for "GRAND TOTAL".
        collapsed = re.sub(r"(.)\1+", r"\1", joined_text)
        if "GRANDTOTAL" in collapsed.replace(" ", ""):
            return sorted(line, key=lambda w: w["x0"])
    return None


def _merge_wrapped_header_words(header_line):
    """Column headers ("Trinidad &\\nTobago") wrap onto a second visual
    line at nearly the same x-position - extract_words() already
    returns them on one logical text line here since Markets/Algeria/...
    all shared one `top` cluster in every inspected sample, so this is
    a passthrough kept for clarity/robustness rather than doing real
    merging work."""
    return [w for w in header_line if w["text"] not in ("Markets",)]


def build_year_row(export_totals, spot_totals):
    """Combine the two tables into one wide row: for every country/
    basin label that appears in EITHER table, Total/Spot/Contracted -
    matched by GIIGNL's own label spelling, which is consistent between
    a report edition's two tables (not guaranteed consistent ACROSS
    edition years - see the module docstring)."""
    all_labels = set(export_totals) | set(spot_totals)
    row = {}
    for label in all_labels:
        total = export_totals.get(label)
        spot = spot_totals.get(label)
        clean_label = re.sub(r"\s+", " ", label).strip()
        if total is not None:
            row[f"{clean_label}_Total_MT"] = total
        if spot is not None:
            row[f"{clean_label}_Spot_ShortTerm_MT"] = spot
        if total is not None and spot is not None:
            row[f"{clean_label}_Contracted_MT"] = total - spot
    return row


def process_year(year):
    import pdfplumber
    from io import BytesIO

    print(f"[{year}] fetching ...", file=sys.stderr)
    content = fetch_pdf_bytes(year)
    if content is None:
        print(f"  no report at the expected URL for {year} - skipping.", file=sys.stderr)
        return None

    with pdfplumber.open(BytesIO(content)) as pdf:
        export_totals = find_export_totals_table(pdf)
        spot_totals = find_spot_shortterm_grand_total(pdf)

    if not export_totals:
        print(f"  WARNING: could not find the export-totals table for {year} - skipping.", file=sys.stderr)
        return None
    if not spot_totals:
        print(f"  WARNING: could not find the spot/short-term grand-total row for {year} - skipping.",
              file=sys.stderr)
        return None

    print(f"  export totals found for {len(export_totals)} countries/basins, "
          f"spot/short-term for {len(spot_totals)}", file=sys.stderr)
    return build_year_row(export_totals, spot_totals)


NOTES_LINES = [
    "UNITS",
    "All volumes in million tonnes (MT) per year. *_Total_MT: total LNG exports for that country/basin "
    "that year. *_Spot_ShortTerm_MT: the portion delivered under contracts of 4 years or less (GIIGNL's "
    "own definition). *_Contracted_MT: Total - Spot/ShortTerm (long-term contracted volume), only "
    "computed where both source numbers were found.",
    "",
    "SCOPE",
    "Exporting countries and basin subtotals (e.g. 'Atlantic Basin1', 'Middle East1', 'Pacific Basin1') "
    "exactly as GIIGNL's own report labels them that year - not a fixed list maintained here, so a "
    "country GIIGNL adds or renames shows up/changes as its own column. One row per report year.",
    "",
    "PARSING CAVEATS",
    "Pulled from GIIGNL's public PDF annual reports (no API or bulk file exists - confirmed via "
    "GIIGNL_DISCOVERY.py) using positional (x-coordinate) word matching, not a plain table-grid reader - "
    "the source PDF's own column-boundary rendering merges some cells when read by a naive table "
    "extractor. A year whose two source tables couldn't both be found is skipped (see the pull's own "
    "log), not silently filled with wrong numbers.",
    "",
    "COVERAGE",
    "Annual only - GIIGNL publishes one report per year, each covering the prior calendar year's trade. "
    "Only years actually found at the known report URL are included.",
    "",
    "SOURCE",
    f"GIIGNL (Groupe International des Importateurs de Gaz Naturel Liquefie) annual 'LNG Industry' "
    f"report, public PDF: {BASE_URL}",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "PARSING CAVEATS", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    rows = {}
    for year in CANDIDATE_YEARS:
        row = process_year(year)
        if row is not None:
            rows[year] = row

    if not rows:
        print("No year's report could be parsed - nothing to save.", file=sys.stderr)
        sys.exit(1)

    df = pd.DataFrame.from_dict(rows, orient="index")
    df.index.name = "report_year"
    df = df.sort_index()
    df = df[sorted(df.columns)]

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": df}, NOTES_LINES, NOTES_SECTION_TITLES)

    print(f"Saved {len(df)} year(s) to {args.out}")
    print(df)


if __name__ == "__main__":
    main()
