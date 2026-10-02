"""
One-off pull of UK (UKCS) oil and gas production and reserves from the
North Sea Transition Authority (NSTA) data centre.

nstauthority.co.uk is blocked from the editing sandbox, so the exact file
names/layouts were not inspectable: the script crawls the NSTA data-centre
pages for production / reserves downloads (.xlsx/.xls/.csv), saves the raw
files, and writes every parsed table to one workbook (raw tabs, Units tab
listing each source URL). A chart sheet is added for each table that has a
date or year column plus numeric columns (annual series show the year,
monthly show mmm/yy). The run log prints the links found and the headers
parsed so layout drift is visible.
"""

import argparse
import io
import os
import re
import sys
from urllib.parse import urljoin

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

import xlsx_charts
import xlsx_notes

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 90)
BASE = "https://www.nstauthority.co.uk"
SEED_PAGES = [
    BASE + "/data-centre/",
    BASE + "/data-centre/data-downloads-and-publications/",
    BASE + "/data-centre/data-downloads-and-publications/production-data/",
    BASE + "/data-centre/data-downloads-and-publications/reserves-and-resources/",
    BASE + "/data-centre/data-downloads-and-publications/production-projections/",
]
KEYWORDS = ("production", "reserve", "resource")
FILE_RE = re.compile(r'href="([^"]+?\.(?:xlsx|xls|csv)(?:\?[^"]*)?)"', re.I)
PAGE_RE = re.compile(r'href="([^"]*data-centre[^"#]*)"', re.I)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "uk_nsta_production_reserves.xlsx")
RAW_DIR = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "uk_nsta_raw")


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r


def discover():
    """Seed pages plus one level of data-centre sub-pages; return {file_url: page}."""
    files, seen, queue = {}, set(), list(SEED_PAGES)
    while queue and len(seen) < 150:
        page = queue.pop(0)
        if page in seen:
            continue
        seen.add(page)
        try:
            html = get(page).text
        except Exception as e:
            print(f"page failed {page}: {e}", file=sys.stderr)
            continue
        for href in FILE_RE.findall(html):
            url = urljoin(page, href)
            files.setdefault(url, page)
        for href in PAGE_RE.findall(html):
            url = urljoin(page, href)
            if url not in seen and url.startswith(BASE) and not FILE_RE.search(f'href="{url}"'):
                queue.append(url)
    print(f"Crawled {len(seen)} pages, found {len(files)} files", file=sys.stderr)
    for p in sorted(seen):
        print("  page:", p, file=sys.stderr)
    for u in files:
        print("  file:", u, file=sys.stderr)
    return files


def parse(url, content):
    """Return {tab_name: DataFrame} for a downloaded file."""
    name = os.path.splitext(os.path.basename(url.split("?")[0]))[0][:20]
    out = {}
    if url.lower().split("?")[0].endswith(".csv"):
        out[name] = pd.read_csv(io.BytesIO(content), low_memory=False, encoding_errors="replace")
    else:
        for sheet, df in pd.read_excel(io.BytesIO(content), sheet_name=None).items():
            out[f"{name}_{sheet}"[:31]] = df
    return {k: v.dropna(how="all").dropna(axis=1, how="all") for k, v in out.items()}


def chart_frame(df):
    """Date/year column + up to 5 summed numeric columns, or None."""
    if len(df) < 3:
        return None
    for col in df.columns[:6]:
        low = str(col).lower()
        if not any(k in low for k in ("date", "month", "year", "period")):
            continue
        key = pd.to_datetime(df[col], errors="coerce") if "year" not in low else pd.to_datetime(
            pd.to_numeric(df[col], errors="coerce").dropna().astype(int).astype(str), format="%Y", errors="coerce")
        num = df.drop(columns=[col]).apply(pd.to_numeric, errors="coerce")
        num = num.loc[:, num.notna().mean() > 0.5]
        if key.notna().sum() < 3 or num.empty:
            continue
        top = num.sum().sort_values(ascending=False).index[:5]
        g = num[top].groupby(key.reindex(num.index)).sum()
        g = g[g.index.notna()].sort_index()
        annual = "year" in low
        return g, ("%Y" if annual else "%b/%y")
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    os.makedirs(RAW_DIR, exist_ok=True)

    sheets, sources, charts = {}, [], []
    for url, page in discover().items():
        try:
            content = get(url).content
            with open(os.path.join(RAW_DIR, os.path.basename(url.split("?")[0])), "wb") as f:
                f.write(content)
            for tab, df in parse(url, content).items():
                print(f"{tab}: shape={df.shape} headers={list(df.columns)[:10]}", file=sys.stderr)
                tab = tab if tab not in sheets else f"{tab[:27]}_{len(sheets)}"
                sheets[tab] = df
                sources.append(f"{tab}: {url}  (listed on {page})")
                cf = chart_frame(df)
                if cf is not None:
                    charts.append((tab, *cf))
        except Exception as e:
            print(f"file failed {url}: {e}", file=sys.stderr)
    if not sheets:
        sys.exit("No NSTA files could be downloaded/parsed")

    notes = ["Units and notes",
             "UK Continental Shelf oil and gas production and reserves, from the North Sea Transition Authority (NSTA) data centre.",
             "One-off pull. Units are as published in each source tab (check the original headers); not converted.",
             "", "Sources"] + sources
    xlsx_notes.write_workbook(args.out, sheets, notes, {"Units and notes", "Sources"})

    from openpyxl import load_workbook
    wb = load_workbook(args.out)
    for tab, g, datefmt in charts:
        xlsx_charts.add_chart_sheet(args.out, g, f"NSTA - {tab}", "As published", kind="line",
                                    sheet_name=f"Chart {tab}"[:31], date_format=datefmt, wb=wb)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Wrote {args.out}: {len(sheets)} data tabs, {len(charts)} charts", file=sys.stderr)


if __name__ == "__main__":
    main()
