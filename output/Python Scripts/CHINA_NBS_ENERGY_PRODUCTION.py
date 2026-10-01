"""
Pull China's National Bureau of Statistics (NBS) monthly "Energy
Production in {Month} {Year}" press release and extract the headline
MONTHLY (not year-to-date, not year-on-year %) output figures: raw
coal, crude oil, crude oil processing volume, natural gas, and
electricity generation.

Found via CHINA_NBS_DISCOVERY.py - no API, no bulk CSV/Excel (the
site's own query engine at data.stats.gov.cn is blocked outright by a
WAF, 403 "reason:UrlACL", for any automated request), but this
specific monthly release is a real, reachable page with a stable
prose template every month, e.g.:

  "In August, the raw coal production by industrial enterprises above
  the designated size was 360 million tons, down by 7.7% year on
  year, ... From January to August, the raw coal production ... was
  3.06 billion tons, a year-on-year decrease of 3.3%."

Only the "In {Month}, ... was X {unit}" sentence is extracted (the
monthly figure) - the "From January to {Month}, ... was Y {unit}"
sentence (year-to-date, in a different unit scale) and the year-on-
year % figures are deliberately not pulled.

DISCOVERY of past editions: the press-release index
(stats.gov.cn/english/PressRelease/) mixes every release type (CPI,
PMI, real estate, ...) across many paginated pages - "Energy
Production" releases are found by scanning that index rather than
guessed at from a URL pattern (each release's own URL has an opaque
numeric ID, not derivable from its date). Pagination itself follows
the pattern embedded in the index page's own JS: page 1 is the base
URL, page N (N>=2) is index_{N-1}.html - confirmed working (page 0 =
Aug 2026, page 10 = Nov 2025), NOT the list_N.html pattern first
guessed (all 404).
"""

import argparse
import os
import re
import sys
import time
from datetime import datetime

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

INDEX_BASE = "https://www.stats.gov.cn/english/PressRelease/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)
FETCH_ATTEMPTS = 3
RETRY_BACKOFF_SECONDS = [10, 30]

# Bounded crawl depth for the index, not an assumption that history
# stops here - just how far back this pull looks for now (roughly
# 1-1.2 releases of every type per page observed, so this reaches
# multiple years back). Increase if a deeper backfill is wanted later.
MAX_INDEX_PAGES = 60
# Stop early only after this many CONSECUTIVE non-200 pages (the index
# itself failing, not just a page with no Energy Production release on
# it - plenty of index pages have zero, since it's a mixed-topic feed).
MAX_CONSECUTIVE_FAILURES = 5

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "china_nbs_energy_production_monthly.xlsx")

TITLE_RE = re.compile(r"Energy Production in (\w+) (\d{4})")

# Each entry: (output column, regex matching the MONTHLY "In {Month},
# ... was X {unit}" sentence - never the "From January to {Month}"
# year-to-date one, which always follows immediately after with a
# different, larger unit scale). group(1)=value, group(2)=unit word.
INDICATOR_PATTERNS = {
    "Raw_Coal_Mt": (
        r"In \w+, the raw coal production by industrial enterprises above the designated size "
        r"was ([\d,]+\.?\d*) (million|billion) tons",
        "tons",
    ),
    "Crude_Oil_Mt": (
        r"In \w+, the crude oil production by industrial enterprises above the designated size "
        r"was ([\d,]+\.?\d*) (million|billion) tons",
        "tons",
    ),
    "Crude_Oil_Processing_Mt": (
        r"In \w+, the processing volume of crude oil by industrial enterprises above the designated size "
        r"was ([\d,]+\.?\d*) (million|billion) tons",
        "tons",
    ),
    "Natural_Gas_Bcm": (
        r"In \w+, the production of natural gas by industrial enterprises above the designated size "
        r"was ([\d,]+\.?\d*) (million|billion) cubic meters",
        "cubic meters",
    ),
    "Electricity_Generation_BkWh": (
        r"In \w+, electricity generation by industrial enterprises above the designated size "
        r"was ([\d,]+\.?\d*) (million|billion) kWh",
        "kWh",
    ),
}

# Canonical unit each output column is normalized to, regardless of
# which scale word ("million"/"billion") the prose happened to use
# that month - confirmed from real editions that coal/oil consistently
# report the monthly figure in "million tons", gas in "billion cubic
# meters", electricity in "billion kWh", but normalizing rather than
# assuming keeps a future edition's differently-scaled wording correct.
CANONICAL_SCALE = {
    "Raw_Coal_Mt": "million",
    "Crude_Oil_Mt": "million",
    "Crude_Oil_Processing_Mt": "million",
    "Natural_Gas_Bcm": "billion",
    "Electricity_Generation_BkWh": "billion",
}


def _fetch(url):
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                time.sleep(RETRY_BACKOFF_SECONDS[attempt - 1])
    raise last_error


def find_energy_production_releases():
    """Scans the press-release index (paginated) for every "Energy
    Production in {Month} {Year}" link, returning [(month, year, url), ...].
    """
    releases = []
    consecutive_failures = 0
    for i in range(MAX_INDEX_PAGES):
        url = INDEX_BASE if i == 0 else f"{INDEX_BASE}index_{i}.html"
        try:
            r = _fetch(url)
        except requests.RequestException:
            r = None
        if r is None or r.status_code != 200:
            consecutive_failures += 1
            print(f"  index page {i}: unreachable ({consecutive_failures} consecutive)", file=sys.stderr)
            if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                print(f"  stopping after {consecutive_failures} consecutive unreachable index pages",
                      file=sys.stderr)
                break
            continue
        consecutive_failures = 0

        for m in re.finditer(r'<a\b([^>]*)>', r.text):
            attrs = m.group(1)
            href_m = re.search(r'href="([^"]*)"', attrs)
            title_m = re.search(r'title="([^"]*)"', attrs)
            if not href_m or not title_m:
                continue
            title_match = TITLE_RE.search(title_m.group(1))
            if not title_match:
                continue
            month, year = title_match.group(1), int(title_match.group(2))
            full_url = href_m.group(1)
            if full_url.startswith("./"):
                full_url = INDEX_BASE + full_url[2:]
            releases.append((month, year, full_url))
    return releases


def parse_release(url):
    r = _fetch(url)
    if r is None:
        return None
    text = re.sub(r"<[^>]+>", " ", r.text)
    text = re.sub(r"\s+", " ", text)

    row = {}
    for column, (pattern, _unit_label) in INDICATOR_PATTERNS.items():
        m = re.search(pattern, text)
        if m is None:
            continue
        value = float(m.group(1).replace(",", ""))
        scale = m.group(2)
        canonical = CANONICAL_SCALE[column]
        if scale != canonical:
            # normalize to the canonical scale (1000x between million/billion)
            value = value * 1000 if canonical == "million" else value / 1000
        row[column] = value
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    print("Scanning press-release index for Energy Production releases ...", file=sys.stderr)
    releases = find_energy_production_releases()
    print(f"  found {len(releases)} releases", file=sys.stderr)

    rows = {}
    for month, year, url in releases:
        try:
            report_date = datetime.strptime(f"{month} {year}", "%B %Y").date()
        except ValueError:
            print(f"  WARNING: could not parse month/year from {month!r} {year!r} - skipping", file=sys.stderr)
            continue
        if report_date in rows:
            continue
        print(f"[{report_date}] fetching {url} ...", file=sys.stderr)
        try:
            row = parse_release(url)
        except requests.RequestException as e:
            print(f"  WARNING: failed to fetch {url}: {e} - skipping", file=sys.stderr)
            continue
        if not row:
            print(f"  WARNING: no indicators matched for {report_date} - skipping", file=sys.stderr)
            continue
        print(f"  matched {len(row)}/{len(INDICATOR_PATTERNS)} indicators", file=sys.stderr)
        rows[report_date] = row

    if not rows:
        print("No release could be parsed - nothing to save.", file=sys.stderr)
        sys.exit(1)

    df = pd.DataFrame.from_dict(rows, orient="index")
    df.index.name = "report_month"
    df = df.sort_index()
    df = df[sorted(df.columns)]

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": df}, NOTES_LINES, NOTES_SECTION_TITLES)

    print(f"Saved {len(df)} month(s) to {args.out}")
    print(df)


NOTES_LINES = [
    "UNITS",
    "Raw_Coal_Mt / Crude_Oil_Mt / Crude_Oil_Processing_Mt: million tonnes for that calendar month. "
    "Natural_Gas_Bcm: billion cubic meters for that month. Electricity_Generation_BkWh: billion kWh for "
    "that month. All figures are MONTHLY only - NBS's release also gives year-to-date and year-on-year %% "
    "figures, deliberately not pulled here.",
    "",
    "SCOPE",
    "Coverage is 'industrial enterprises above the designated size' (NBS's own definition: annual main "
    "business revenue of RMB 20 million or above) - NBS's own stated basis for these figures, not all "
    "national output. Crude_Oil_Processing_Mt is refinery throughput (a separate figure from crude oil "
    "production itself).",
    "",
    "PARSING CAVEATS",
    "Pulled from NBS's own monthly 'Energy Production in {Month} {Year}' English-language press release "
    "(no API or bulk file exists - data.stats.gov.cn's query engine returns 403 UrlACL for any automated "
    "request, confirmed via CHINA_NBS_DISCOVERY.py). Each month's monthly figure is read from its own "
    "'In {Month}, ... was X {unit}' sentence specifically (never the immediately-following year-to-date "
    "sentence, which uses a different, larger unit scale). A month whose release couldn't be found or "
    "whose sentence didn't match is left out entirely (see the pull's own log), not filled with a wrong "
    "number.",
    "",
    "COVERAGE",
    "Only as far back as the press-release index was actually crawled this run (see the pull's own log for "
    "how many releases were found) - not assumed to include the full history NBS may have published.",
    "",
    "SOURCE",
    f"National Bureau of Statistics of China, monthly 'Energy Production' press release, English section: "
    f"{INDEX_BASE}",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "PARSING CAVEATS", "COVERAGE", "SOURCE"}


if __name__ == "__main__":
    main()
