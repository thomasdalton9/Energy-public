"""
Follow-up to EIA_COAL_DISCOVERY.py, which found EIA's structured API
(api.eia.gov) has no weekly/current coal benchmark price - only an
annual, ~2-year-stale series. The user pointed at
https://www.eia.gov/coal/markets/#tabs-prices-2 directly - EIA's own
"Coal Markets" report page, which does show current weekly benchmark
prices (Central Appalachia, Northern Appalachia, Illinois Basin, Powder
River Basin, Uinta Basin) as a narrative report/chart, per general
knowledge of this report - not yet confirmed against the live page or
checked for a download link or a data file/API behind it.

This fetches the live page and:
  1. Saves the raw HTML for inspection.
  2. Looks for direct download links (.csv/.xlsx/.xls/.json).
  3. Looks for embedded JSON/data blobs the page's own JS might read
     from (common in EIA's newer report pages - a <script
     type="application/json"> block or a fetch() call to a data file).
  4. Prints any table content found server-side (vs. client-rendered).
"""

import re
import sys

import requests

URL = "https://www.eia.gov/coal/markets/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)


def main():
    print(f"Fetching {URL} ...", file=sys.stderr)
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    print(f"  status {r.status_code}, {len(r.content):,} bytes", file=sys.stderr)
    r.raise_for_status()
    html = r.text

    with open("/tmp/eia_coal_markets.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("\n=== Direct file links (.csv/.xlsx/.xls/.json/.pdf) ===")
    links = re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.IGNORECASE)
    file_links = [l for l in links if re.search(r"\.(csv|xlsx?|json|pdf)(\?|$)", l, re.IGNORECASE)]
    for l in sorted(set(file_links)):
        print(f"  {l}")
    if not file_links:
        print("  none found")

    print("\n=== fetch()/XHR calls to *.json or /api/ in inline scripts ===")
    fetch_calls = re.findall(r'fetch\((["\'][^"\']+["\'])', html)
    api_like = re.findall(r'["\']([^"\']*(?:/api/|\.json)[^"\']*)["\']', html)
    for c in sorted(set(fetch_calls)):
        print(f"  fetch: {c}")
    for a in sorted(set(api_like))[:30]:
        print(f"  api-like: {a}")

    print("\n=== <script type=\"application/json\"> blocks ===")
    json_blocks = re.findall(r'<script[^>]+type=["\']application/json["\'][^>]*>(.*?)</script>', html, flags=re.DOTALL)
    print(f"  found {len(json_blocks)} block(s)")
    for b in json_blocks[:2]:
        print(f"  preview: {b[:500]}")

    print("\n=== Table tags present server-side ===")
    print(f"  <table count: {html.lower().count('<table')}, <tr count: {html.lower().count('<tr')}")

    print("\n=== Keyword check: basin names / 'coal markets' data hints ===")
    for kw in ["Central Appalachia", "Northern Appalachia", "Illinois Basin", "Powder River",
               "Uinta", "dollars per short ton", "spot price"]:
        print(f"  {kw!r}: {'FOUND' if kw.lower() in html.lower() else 'not found'}")

    print(f"\nSaved full HTML to /tmp/eia_coal_markets.html ({len(html):,} chars) - "
          "not committed to the repo (too large/noisy), just for this run's log.")


if __name__ == "__main__":
    main()
