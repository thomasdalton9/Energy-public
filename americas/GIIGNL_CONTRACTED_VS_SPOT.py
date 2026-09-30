"""
Pull GIIGNL's annual LNG Industry report PDFs and extract, per major
LNG-exporting country/basin and per report year, contracted (long-term)
vs spot/short-term export volumes.

Data source: GIIGNL (Groupe International des Importateurs de Gaz
Naturel Liquefie) annual reports, published as public PDFs at
https://giignl-documents.s3.fr-par.scw.cloud/public/ar-{year}-annual-report.pdf
Found via GIIGNL_DISCOVERY.py / GIIGNL_PDF_INSPECT.py - no API, no CSV/
Excel, PDF only, but with genuinely real (non-scanned) extractable text
and two data tables that together answer contracted-vs-spot:

  1. Total exports by exporting country/basin - a normal, cleanly-
     extractable table ("ATLANTIC BASIN1 / MIDDLE EAST1 / PACIFIC
     BASIN1" sections, each row Country, MT, Global Share, Var Mt,
     Var %).
  2. A "Spot and Short-Term* LNG Quantities (in MT) received in {year}"
     bilateral matrix (importing market x exporting country/basin) -
     footnoted "*Quantities delivered under contracts of a duration of
     4 years or less". Its own GRAND TOTAL row, read per exporting
     country/basin column, is that country's total spot+short-term
     EXPORT volume for the year.

Contracted (long-term) = Total exports (table 1) - Spot/short-term
exports (table 2's grand total row), per exporting country/basin.

FIRST ATTEMPT (kept only in git history) tried to auto-discover every
column generically via raw word-position clustering, on the theory
that not hardcoding a country list is more maintainable. In practice
it made things worse: individual words from multi-word headers
("Trinidad & Tobago") got treated as separate columns, and rows from
an unrelated IMPORT-side table on the same page (Belgium, France,
Japan...) got swept in alongside real export rows, because nothing
distinguished "a row that looks like this table's layout" from "a row
that's actually in this table". GIIGNL's own list of major LNG
exporters is stable, well-known industry knowledge (unlike, say,
EIA-930's fuel-type categories, which genuinely do change) - so this
version matches against MAJOR_EXPORTERS by name instead of trying to
discover columns generically. A real exporter GIIGNL adds that isn't
in this list yet simply won't show up until added here.

PARSING APPROACH:
  - Table 1: read via pdfplumber's normal extract_tables(). A table is
    only trusted for exporter matches if it also contains one of
    GIIGNL's own basin-subtotal rows (Atlantic Basin1/Middle East1/
    Pacific Basin1) - name-matching alone isn't enough, since the US,
    Canada, Mexico and Egypt (among others) also show up as small
    volumes in an unrelated IMPORT-side table elsewhere in the same
    report (confirmed the hard way: an earlier version of this script
    silently took the US's ~0.4 MT import row instead of its real
    ~109 MT export total, because it matched on name alone and took
    whichever table it hit first).
  - Table 2: the header row extracts cleanly as one list via
    extract_tables() (each multi-word country name stays one cell,
    unlike raw word-clustering), so header cell text AND that row's
    pdfplumber-reported cell bounding boxes are used to get each known
    exporter's column x-range directly from the real column boundaries.
    The GRAND TOTAL row's own line is then found by text search, and
    each exporter's value is read by cropping the page to (that
    column's x-range) x (that row's y-range) and extracting text from
    just that crop - exact positional alignment, not word-matching
    heuristics, and a blank/zero cell just yields an empty crop (kept
    as NaN) instead of shifting every later column.

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

# Exactly as GIIGNL itself spells these in both source tables (confirmed
# against the 2026-edition PDF) - a real exporter missing from this list
# (new project, or one this script's author didn't see reported yet)
# simply won't get a column until added here. Basin subtotals included
# since they're genuinely useful totals, not because they're needed to
# find the table (table 1's own filtering no longer depends on them).
MAJOR_EXPORTERS = [
    "Algeria", "Angola", "Cameroon", "Congo", "Egypt", "Equatorial Guinea", "Mauritania",
    "Mexico", "Nigeria", "Norway", "Russia Europe", "Trinidad & Tobago", "USA",
    "Oman", "Qatar", "UAE",
    "Australia", "Brunei", "Canada", "Indonesia", "Malaysia", "Mozambique",
    "Papua New Guinea", "Peru", "Russia Asia",
    "Atlantic Basin1", "Middle East1", "Pacific Basin1",
]

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "giignl_contracted_vs_spot_annual.xlsx")


def _normalize(s):
    return re.sub(r"\s+", " ", str(s)).strip().casefold()


_EXPORTER_LOOKUP = {_normalize(name): name for name in MAJOR_EXPORTERS}


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


_BASIN_SUBTOTALS = {"Atlantic Basin1", "Middle East1", "Pacific Basin1"}


def find_export_totals(pdf):
    """Table 1: total exports by exporting country/basin. Matching by
    name alone isn't enough - the US, Canada, Mexico and Egypt (among
    others) also show up as small volumes in an unrelated IMPORT-side
    table elsewhere in the same report (confirmed by an early version
    of this script silently picking up the US's 0.4 MT import row
    instead of its real ~109 MT export total). A table is only trusted
    for exporter matches if it also contains one of GIIGNL's own basin-
    subtotal rows (Atlantic Basin1/Middle East1/Pacific Basin1) -
    a label distinctive enough to only appear in the real export-totals
    table, unlike a plain country name."""
    result = {}
    for page in pdf.pages:
        for table in page.extract_tables():
            rows = [r for r in table if r and r[0]]
            is_export_table = any(
                _EXPORTER_LOOKUP.get(_normalize(r[0])) in _BASIN_SUBTOTALS for r in rows
            )
            if not is_export_table:
                continue
            for row in rows:
                canonical = _EXPORTER_LOOKUP.get(_normalize(row[0]))
                if canonical is None or canonical in result:
                    continue
                if len(row) < 2 or row[1] in (None, ""):
                    continue
                try:
                    mt = float(str(row[1]).replace(",", ""))
                except ValueError:
                    continue
                result[canonical] = mt
    return result


def find_spot_shortterm_totals(pdf):
    """Table 2's GRAND TOTAL row, read per known exporter via exact
    column-boundary cropping (see module docstring) rather than word-
    position guessing."""
    for page in pdf.pages:
        text = page.extract_text() or ""
        if "Spot and Short" not in text and "SPOT AND SHORT" not in text.upper():
            continue
        collapsed = _collapse_doubled_letters(text.upper().replace("\n", " "))
        if "GRAND TOTAL" not in collapsed:
            continue  # some editions may wrap/omit the title differently

        header_cells = _find_header_row_cells(page)
        if header_cells is None:
            continue
        total_row_bbox = _find_grand_total_row_bbox(page)
        if total_row_bbox is None:
            continue
        total_top, total_bottom = total_row_bbox

        result = {}
        for header_text, (x0, x1) in header_cells.items():
            canonical = _EXPORTER_LOOKUP.get(_normalize(header_text))
            if canonical is None or canonical in result:
                continue
            crop = page.within_bbox((x0, total_top, x1, total_bottom))
            cell_text = (crop.extract_text() or "").strip()
            if not cell_text:
                continue
            match = re.search(r"-?\d+(\.\d+)?", cell_text.replace(",", ""))
            if match:
                result[canonical] = float(match.group())
        return result if result else None
    return None


def _find_header_row_cells(page):
    """Uses find_tables() to get the spot/short-term matrix's own
    detected table object, so the header row's cell x-ranges are the
    PDF's real column boundaries - not re-derived from word positions."""
    for table in page.find_tables():
        extracted = table.extract()
        if not extracted or not extracted[0]:
            continue
        header_row_text = extracted[0]
        if "Markets" not in header_row_text:
            continue
        header_cells = table.rows[0].cells
        result = {}
        for text, cell in zip(header_row_text, header_cells):
            if text and cell is not None:
                x0, _, x1, _ = cell
                result[text] = (x0, x1)
        return result
    return None


def _find_grand_total_row_bbox(page):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    lines = _group_words_into_lines(words)
    for line in lines:
        joined = _collapse_doubled_letters("".join(w["text"] for w in line).upper())
        if "GRANDTOTAL" in joined.replace(" ", ""):
            top = min(w["top"] for w in line)
            bottom = max(w["bottom"] for w in line)
            return (top, bottom)
    return None


def _collapse_doubled_letters(s):
    """GIIGNL's PDF renders certain bold/spaced headings with every
    letter doubled (a font/kerning artifact, not a typo) -
    "GGRRAANNDD TTOOTTAALL" for "GRAND TOTAL", "AASSIIAA" for "ASIA"."""
    return re.sub(r"(.)\1+", r"\1", s)


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


def build_year_row(export_totals, spot_totals):
    row = {}
    for name in MAJOR_EXPORTERS:
        total = export_totals.get(name)
        spot = spot_totals.get(name)
        if total is not None:
            row[f"{name}_Total_MT"] = total
        if spot is not None:
            row[f"{name}_Spot_ShortTerm_MT"] = spot
        if total is not None and spot is not None:
            row[f"{name}_Contracted_MT"] = total - spot
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
        export_totals = find_export_totals(pdf)
        spot_totals = find_spot_shortterm_totals(pdf)

    if not export_totals:
        print(f"  WARNING: could not find the export-totals table for {year} - skipping.", file=sys.stderr)
        return None
    if not spot_totals:
        print(f"  WARNING: could not find the spot/short-term grand-total row for {year} - skipping.",
              file=sys.stderr)
        return None

    print(f"  export totals matched for {len(export_totals)}/{len(MAJOR_EXPORTERS)} known exporters, "
          f"spot/short-term for {len(spot_totals)}/{len(MAJOR_EXPORTERS)}", file=sys.stderr)
    return build_year_row(export_totals, spot_totals)


NOTES_LINES = [
    "UNITS",
    "All volumes in million tonnes (MT) per year. *_Total_MT: total LNG exports for that country/basin "
    "that year. *_Spot_ShortTerm_MT: the portion delivered under contracts of 4 years or less (GIIGNL's "
    "own definition). *_Contracted_MT: Total - Spot/ShortTerm (long-term contracted volume), only "
    "computed where both source numbers were found for that country/year.",
    "",
    "SCOPE",
    "A fixed list of major LNG exporters (see MAJOR_EXPORTERS in the script) plus GIIGNL's own basin "
    "subtotals (Atlantic Basin1, Middle East1, Pacific Basin1) - matched by name against GIIGNL's own "
    "report tables, not auto-discovered. A country missing from a given year's columns means GIIGNL's "
    "report that year didn't report that country in one or both source tables (e.g. too small to list "
    "individually), not that it exported zero - left blank (NaN) rather than assumed to be 0.",
    "",
    "PARSING CAVEATS",
    "Pulled from GIIGNL's public PDF annual reports (no API or bulk file exists - confirmed via "
    "GIIGNL_DISCOVERY.py). The spot/short-term table's GRAND TOTAL row is read by cropping the page to "
    "each known exporter's own column x-range (from the table's real detected column boundaries) and "
    "that row's y-range, then reading whatever number is in that crop - exact positional alignment, not "
    "a naive table-grid reader (which merges some of this specific table's data-row cells when read "
    "directly). A year whose two source tables couldn't both be found is skipped (see the pull's own "
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
