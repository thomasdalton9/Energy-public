"""
Checks how GIIGNL's report format differs in 2020-2024 editions vs the
2025/2026 editions GIIGNL_CONTRACTED_VS_SPOT.py was built against -
whether the same two tables (basin-anchored export totals, "Spot and
Short-Term" bilateral matrix with a GRAND TOTAL row) exist at all, and
if so, what their header/label spelling actually is that year (not
assumed to match 2025/2026 exactly).
"""

import re
import sys

import requests

BASE_URL = "https://giignl-documents.s3.fr-par.scw.cloud/public/ar-{year}-annual-report.pdf"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
YEARS_TO_CHECK = [2020, 2021, 2022, 2023, 2024]


def collapse_doubled_letters(s):
    return re.sub(r"(.)\1+", r"\1", s)


def inspect_year(year):
    import pdfplumber
    from io import BytesIO

    url = BASE_URL.format(year=year)
    print(f"\n{'=' * 70}\n{year}: {url}\n{'=' * 70}")
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    if r.status_code == 404:
        print("  404 - no report at this URL for this year.")
        return
    r.raise_for_status()
    print(f"  {len(r.content):,} bytes")

    with pdfplumber.open(BytesIO(r.content)) as pdf:
        print(f"  {len(pdf.pages)} pages")
        spot_pages = []
        basin_pages = []
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            upper = text.upper()
            if "SPOT AND SHORT" in upper or "SPOT & SHORT" in upper:
                spot_pages.append(i)
            collapsed = collapse_doubled_letters(upper.replace("\n", " "))
            if "GRAND TOTAL" in collapsed:
                basin_pages.append(("grand_total", i))
            if any(b in upper for b in ["ATLANTIC BASIN", "MIDDLE EAST", "PACIFIC BASIN"]):
                basin_pages.append(("basin_label", i))

        print(f"  pages mentioning 'spot and short-term': {spot_pages}")
        print(f"  pages with grand-total/basin-label markers: {basin_pages}")

        for i in spot_pages[:2]:
            page = pdf.pages[i]
            text = page.extract_text() or ""
            print(f"\n  --- page {i} preview (first 1000 chars) ---")
            print("  " + text[:1000].replace("\n", "\n  "))
            tables = page.extract_tables()
            print(f"  --- page {i} has {len(tables)} table(s); first table's header row (if any): ---")
            if tables and tables[0]:
                print(f"    {tables[0][0]}")


def main():
    for year in YEARS_TO_CHECK:
        try:
            inspect_year(year)
        except Exception as e:
            print(f"  ERROR inspecting {year}: {type(e).__name__}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
