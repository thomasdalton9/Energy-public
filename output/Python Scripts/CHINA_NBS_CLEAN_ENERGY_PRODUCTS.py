"""
Pull China's National Bureau of Statistics (NBS) monthly "Industrial
Production Operation in {Month} {Year}" press release and extract the
headline MONTHLY output figures for:
  - Solar panels (Solar Cells/Photovoltaic Cells) and EVs (New Energy
    Vehicles, plus total Motor Vehicles) - real manufactured-product
    line items.
  - Electricity GENERATION by source in GWh (thermal, hydro, nuclear,
    wind, solar, and the total) - the closest available NBS proxy for
    "wind turbines"/"solar panels" as an output metric where no
    manufactured-equipment figure exists (NBS does not publish wind
    turbine manufacturing output as a headline monthly indicator -
    only generation in kWh).
  - Coal, coke, cement, and natural gas output - everything "coal or
    gas adjacent" from the release's full product list, per explicit
    scope. Cement itself isn't made from coal or gas (its raw material
    is limestone) but is included as a major coal-DEMAND sector -
    Chinese cement kilns are overwhelmingly coal-fired.

Heat pumps are not reported anywhere in this release at all; no NBS
source for them was found (see CHINA_NBS_DISCOVERY.py).

Found via CHINA_NBS_DISCOVERY.py: unlike the "Energy Production"
release (free-form prose), this release's product-output section is a
genuine HTML table - each row is [product name (unit), monthly value,
monthly year-on-year %, year-to-date value, year-to-date year-on-year
%] - confirmed by cross-checking Crude Oil's row here (1843, in
10,000-ton units) against the Energy Production release's prose figure
for the same month (18.43 million tons) - same number, consistent
column order. Only the MONTHLY value (first of the four) is pulled,
matching the same scope decision as CHINA_NBS_ENERGY_PRODUCTION.py.

Historical editions are found the same way as that script: crawling
the press-release index's own pagination (page 1 = base URL, page N =
index_{N-1}.html) for "Industrial Production Operation in {Month}
{Year}" links, since each release's own URL has an opaque numeric ID.
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

MAX_INDEX_PAGES = 60
MAX_CONSECUTIVE_FAILURES = 5

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "china_nbs_clean_energy_products_monthly.xlsx")

TITLE_RE = re.compile(r"Industrial Production Operation in (\w+) (\d{4})")

# Each product's exact label text as it appears in the table (matched
# by substring, since some rows have a "Of which: " prefix) -> output
# column name. Values are converted at parse time from the source's
# own unit to the CONVERSION_FACTORS unit below wherever a factor is
# given (default 1.0, i.e. kept as the source's own unit).
PRODUCT_LABELS = {
    "Solar Cells (Photovoltaic Cells)": "Solar_Cells_10k_kW",
    "New Energy Vehicles": "New_Energy_Vehicles_10k_units",
    "Motor Vehicles": "Motor_Vehicles_10k_units",
    "Electricity Generation of Industrial Enterprises Above the Designated Size": "Total_Generation_GWh",
    "Thermal Power": "Thermal_Generation_GWh",
    "Hydropower": "Hydro_Generation_GWh",
    "Nuclear Power": "Nuclear_Generation_GWh",
    "Wind Power": "Wind_Generation_GWh",
    "Solar Power": "Solar_Generation_GWh",
    "Coal (10,000 tons)": "Coal_Mt",
    "Coke (10,000 tons)": "Coke_Mt",
    "Natural Gas": "Natural_Gas_Bcm",
    # Not a coal/gas OUTPUT itself - included because Chinese cement
    # kilns are overwhelmingly coal-fired (a major coal-DEMAND sector),
    # per the user's explicit "anything coal or gas adjacent" scope.
    "Cement (10,000 tons)": "Cement_Mt",
}

# Source unit -> output unit, per column (columns not listed here are
# kept as-is, factor 1.0):
#   - generation-by-source: source's "100 million kWh" -> GWh
#     (100,000,000 kWh = 100 GWh, since 1 GWh = 1,000,000 kWh)
#   - Coal/Coke: source's "10,000 tons" -> million tonnes (Mt)
#     (10,000 tons / 1,000,000 = 0.01 Mt per source unit... i.e. x0.01)
#   - Natural Gas: source's "100 million cubic meters" -> billion
#     cubic meters (Bcm) (100M m3 = 0.1 Bcm)
# Cross-checked against CHINA_NBS_ENERGY_PRODUCTION.py's prose-parsed
# figures for the same month: Natural_Gas_Bcm 21.4 here (214 x 0.1)
# matches that script's 21.4 exactly.
CONVERSION_FACTORS = {
    "Total_Generation_GWh": 100,
    "Thermal_Generation_GWh": 100,
    "Hydro_Generation_GWh": 100,
    "Nuclear_Generation_GWh": 100,
    "Wind_Generation_GWh": 100,
    "Solar_Generation_GWh": 100,
    "Coal_Mt": 0.01,
    "Coke_Mt": 0.01,
    "Cement_Mt": 0.01,
    "Natural_Gas_Bcm": 0.1,
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


def find_industrial_production_releases():
    """Scans the press-release index (paginated) for every "Industrial
    Production Operation in {Month} {Year}" link, returning
    [(month, year, url), ...].
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
    row = {}
    for tr_match in re.finditer(r"<tr[^>]*>(.*?)</tr>", r.text, re.S):
        spans = re.findall(r"<span[^>]*>([^<]+)</span>", tr_match.group(1))
        if len(spans) < 2:
            continue
        label = spans[0]
        for product_label, column in PRODUCT_LABELS.items():
            if product_label in label and column not in row:
                try:
                    value = float(spans[1].replace(",", ""))
                except ValueError:
                    continue
                value *= CONVERSION_FACTORS.get(column, 1.0)
                row[column] = value
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    print("Scanning press-release index for Industrial Production Operation releases ...", file=sys.stderr)
    releases = find_industrial_production_releases()
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
            print(f"  WARNING: no products matched for {report_date} - skipping", file=sys.stderr)
            continue
        print(f"  matched {len(row)}/{len(PRODUCT_LABELS)} products", file=sys.stderr)
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
    "Solar_Cells_10k_kW / Motor_Vehicles_10k_units / New_Energy_Vehicles_10k_units: 10,000 kW of "
    "manufacturing capacity (solar cells) or 10,000 units produced (vehicles) that MONTH. "
    "*_Generation_GWh (Total/Thermal/Hydro/Nuclear/Wind/Solar): GWh of electricity GENERATED that month by "
    "that source (not equipment manufactured for Wind/Solar - see SCOPE), converted from the source's own "
    "'100 million kWh' unit (x100). Coal_Mt / Coke_Mt / Cement_Mt: million tonnes for that month, converted "
    "from the source's own '10,000 tons' unit (x0.01). Natural_Gas_Bcm: billion cubic meters for that "
    "month, converted from the source's own '100 million cubic meters' unit (x0.1). All figures are "
    "MONTHLY only - the source release also gives year-to-date and year-on-year %% figures, deliberately "
    "not pulled here.",
    "",
    "SCOPE",
    "Coverage is 'industrial enterprises above the designated size' (NBS's own definition: annual main "
    "business revenue of RMB 20 million or above). New_Energy_Vehicles is a sub-total of Motor_Vehicles "
    "(labelled 'Of which:' in the source); the five generation-by-source columns should sum to "
    "approximately Total_Generation_GWh (small rounding differences from the source's own figures are "
    "expected). Wind turbines are NOT tracked as a manufactured-product line item anywhere in this release "
    "- Wind_Generation_GWh is electricity output, the closest available NBS proxy, not a turbine "
    "production figure. Cement_Mt is not itself a coal or gas product (its raw material is limestone, "
    "calcined to make clinker) - included because Chinese cement kilns are overwhelmingly coal-fired, "
    "making cement output a proxy for coal demand rather than coal/gas supply. Heat pumps are not reported "
    "in this release at all; no NBS monthly source for heat pump output was found (see "
    "CHINA_NBS_DISCOVERY.py) - would need a different (non-NBS) source.",
    "",
    "PARSING CAVEATS",
    "Pulled from NBS's own monthly 'Industrial Production Operation in {Month} {Year}' English-language "
    "press release (no API or bulk file exists - data.stats.gov.cn's query engine returns 403 UrlACL for "
    "any automated request). Each product's row in the release's own HTML table gives 4 values in a fixed "
    "order - [monthly value, monthly year-on-year %, year-to-date value, year-to-date year-on-year %] - "
    "confirmed against the same month's Crude Oil figure in the separate Energy Production release "
    "(1843 in this table's 10,000-ton units = 18.43 million tons there) and against that release's total "
    "electricity figure (9438 in this table's 100-million-kWh units = 943.8 billion kWh there). Only the "
    "first (monthly) value is kept. A month whose release couldn't be found or whose row didn't match is "
    "left out entirely (see the pull's own log), not filled with a wrong number.",
    "",
    "COVERAGE",
    "Only as far back as the press-release index was actually crawled this run (see the pull's own log for "
    "how many releases were found) - not assumed to include the full history NBS may have published.",
    "",
    "SOURCE",
    f"National Bureau of Statistics of China, monthly 'Industrial Production Operation' press release, "
    f"English section: {INDEX_BASE}",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "PARSING CAVEATS", "COVERAGE", "SOURCE"}


if __name__ == "__main__":
    main()
