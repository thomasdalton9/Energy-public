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
from datetime import date, datetime

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


CATEGORY_URL = "https://www.energy.gov.tt/category/publications/energy-industry-bulletins/"
DATA_START = pd.Timestamp("2021-01-01")


def publication_key(url):
    """Upload folder (wp-content/uploads/YYYY/MM) then the publish date
    some file names end with, so an amended re-issue sorts after the
    original."""
    m = re.search(r"/uploads/(\d{4})/(\d{2})/", url)
    d = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})\.xlsx?$", url)
    return ((m.group(1) + m.group(2)) if m else "000000") + (
        f"{d.group(3)}{int(d.group(2)):02d}{int(d.group(1)):02d}" if d else "00000000")


def bulletin_year(url):
    """The year a bulletin covers, from its file name ("January-December-
    2021", "January-May-2026"), not the amendment date some names end with."""
    name = url.rsplit("/", 1)[-1]
    m = re.search(r"January-(?:[A-Za-z]+-)?(20\d\d)", name, re.I) or re.search(r"(20\d\d)", name)
    return int(m.group(1)) if m else None


def list_bulletin_urls():
    """Every bulletin workbook linked from the bulletin category (one
    post per year; full-year editions for past years). Picks up a new
    year's edition (e.g. Jan-Dec 2025) automatically once it's posted."""
    posts = set()
    for page in range(1, 10):
        url = CATEGORY_URL if page == 1 else f"{CATEGORY_URL}page/{page}/"
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            break
        posts |= set(re.findall(r'href="(https://www\.energy\.gov\.tt/[a-z0-9\-]*bulletin[a-z0-9\-]*/)"', r.text))
    # Some yearly posts (e.g. the Jan-Dec 2025 edition) aren't listed in the
    # category pages; the WordPress post sitemap lists every post.
    try:
        r = requests.get("https://www.energy.gov.tt/wp-sitemap-posts-post-1.xml", headers=HEADERS, timeout=TIMEOUT)
        if r.status_code == 200:
            posts |= {u for u in re.findall(r"<loc>([^<]+)</loc>", r.text) if re.search(r"bulletin", u, re.I)}
    except requests.RequestException as e:
        print(f"  sitemap unavailable: {e}", flush=True)
    urls = set()
    for post in sorted(posts):
        years = [int(y) for y in re.findall(r"(20\d\d)", post)]
        if years and max(years) < DATA_START.year:
            continue
        r = requests.get(post, headers=HEADERS, timeout=TIMEOUT)
        for link in re.findall(r'href="([^"]+\.xlsx?)"', r.text, re.I):
            if re.search(r"Consolidated-Monthly-Bulletin", link, re.I):
                urls.add(link.replace("http://", "https://"))
    return sorted(urls)


def read_bulletin(bulletin_url):
    r = requests.get(bulletin_url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(r.content))
    full = xl.parse(find_gas_sheet(xl), header=None)
    production_table, utilization_table = find_tables(full)
    return table_to_long(production_table, "company"), table_to_long(utilization_table, "sector")


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
    # openpyxl hands back date-formatted cells as plain datetime.datetime
    # (or datetime.date), NOT pd.Timestamp - isinstance(val, pd.Timestamp)
    # alone silently matched nothing, producing zero output rows (caught
    # live: the first automated run committed an empty archive).
    month_cols = [i for i, val in enumerate(header) if isinstance(val, (pd.Timestamp, datetime, date))]
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
            # Named month_date, not date - assigning to a name also used
            # as the imported datetime.date type anywhere in this function
            # makes Python treat every reference to that name as the local
            # variable for the WHOLE function, including the isinstance()
            # check above that runs first - caught live as a NameError
            # ("free variable... not associated with a value") the first
            # time this ran against a real bulletin.
            month_date = pd.Timestamp(header.iloc[i]).replace(day=1)
            rows.append({"date": month_date, entity_col: entity, "mmscfd": value, "country": COUNTRY})
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--full", action="store_true", help="re-read every bulletin from 2021, ignoring the archive")
    parser.add_argument("--url", default=None, help="override the bulletin .xlsx URL (skips auto-discovery)")
    args = parser.parse_args()

    production_parts, utilization_parts = [], []
    complete_years = set()
    if os.path.exists(args.out) and not args.full and not args.url:
        # Incremental: keep the archive, and only read bulletins for years
        # that aren't complete in it, plus the current year-to-date one.
        for sheet, parts in (("Production by company", production_parts),
                             ("Utilization by sector", utilization_parts)):
            old = pd.read_excel(args.out, sheet_name=sheet)
            old = old.drop(columns=[c for c in old.columns if str(c).startswith("Unnamed")])
            old["date"] = pd.to_datetime(old["date"])
            if "source" not in old.columns:
                old["source"] = ""
            old["source"] = old["source"].fillna("").astype(str)
            old["pub"] = old["source"].map(publication_key)
            parts.append(old)
        u_old = utilization_parts[0]
        per_year = u_old.groupby(u_old["date"].dt.year)["date"].nunique()
        complete_years = {int(y) for y, n in per_year.items() if n == 12}
        print(f"archive: {u_old['date'].nunique()} months; complete years {sorted(complete_years)}", flush=True)

    if args.url:
        urls = [args.url]
    else:
        current = resolve_current_bulletin_url()
        urls = [u for u in list_bulletin_urls() if bulletin_year(u) not in complete_years]
        if current not in urls:
            urls.append(current)
    print(f"{len(urls)} bulletin workbooks to read", flush=True)

    for bulletin_url in sorted(urls, key=publication_key):
        try:
            p, u = read_bulletin(bulletin_url)
        except Exception as e:
            print(f"  SKIP {bulletin_url}: {type(e).__name__}: {str(e)[:150]}", flush=True)
            continue
        if args.url is None:
            p = p[p["date"] >= DATA_START] if not p.empty else p
            u = u[u["date"] >= DATA_START] if not u.empty else u
        if p.empty and u.empty:
            continue
        print(f"  {bulletin_url.rsplit('/', 1)[-1]}: months "
              f"{u['date'].min():%Y-%m}..{u['date'].max():%Y-%m}", flush=True)
        production_parts.append(p.assign(source=bulletin_url, pub=publication_key(bulletin_url)))
        utilization_parts.append(u.assign(source=bulletin_url, pub=publication_key(bulletin_url)))

    def splice(parts, key):
        """Where two bulletins cover the same month, keep every row of
        that month from the most recently published one (amended
        editions supersede the originals)."""
        df = pd.concat(parts, ignore_index=True)
        newest = df.groupby("date")["pub"].transform("max")
        df = df[df["pub"] == newest].drop(columns="pub")
        # same edition already in the archive and read again: fresh read wins
        df = df.drop_duplicates(["date", key], keep="last")
        return df.sort_values(["date", key]).reset_index(drop=True)

    production = splice(production_parts, "company")
    utilization = splice(utilization_parts, "sector")
    months = pd.date_range(utilization["date"].min(), utilization["date"].max(), freq="MS")
    gaps = [f"{m:%Y-%m}" for m in months if m not in set(utilization["date"])]
    print(f"Spliced {utilization['date'].nunique()} months "
          f"{utilization['date'].min():%Y-%m}..{utilization['date'].max():%Y-%m}; missing: {gaps or 'none'}",
          flush=True)

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
        "Monthly from January 2021. Each year comes from that year's bulletin workbook (full-year editions for "
        "past years, amended Nov 2025; the current year-to-date edition for this year), found on every run "
        "from the ministry's bulletin category pages, so a newly posted year is picked up automatically. Where "
        "two editions cover the same month the most recently published one wins; the 'source' column names the "
        "workbook used. Months not yet reached in a year-to-date edition show as 0 in the source and are "
        "dropped rather than kept as fake zero readings. Bulletin posts are found from both the category pages "
        "and the site's post sitemap (the Jan-Dec 2025 edition is only in the latter). Runs are incremental: years already complete in this "
        "archive are not downloaded again; each run reads the current year-to-date bulletin plus any year with "
        "missing months. Use --full to rebuild from scratch.",
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
