"""
7 of this repo's EU gas-demand scripts currently fall back to ENTSOG's
generic aggregatedData endpoint (adjacent-system categories only, not a
real end-use sector split) because their national TSO wasn't checked
live when they were built - each docstring says so explicitly ("no live
API access to confirm", or for Netherlands "ENTSOG doesn't offer a
finer split" without confirming the TSO itself was actually checked).
Greece is excluded here - its TSO (DESFA) was already tried and found
to have nothing usable.

This checks each remaining country's real national TSO transparency/
open-data page directly, looking for a residential/commercial/power/
industry (or equivalent) sector breakdown that would be better than
ENTSOG's adjacent-system view:

  Denmark     Energinet - energidataservice.dk (known open data API)
  Romania     Transgaz - transgaz.ro
  Portugal    REN (Redes Energeticas Nacionais) - ren.pt
  Belgium     Fluxys - fluxys.com
  Poland      GAZ-SYSTEM - gaz-system.pl
  Italy       Snam Rete Gas - snam.it
  Netherlands GTS (Gasunie Transport Services) - gasunietransportservices.nl

Usage: python3 EU_TSO_SECTOR_DISCOVERY.py
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 45)

SECTOR_KEYWORDS = ["residential", "commercial", "industr", "power", "household", "sector",
                   "huishoud", "industrie", "residenz", "industria", "settore", "gospodarst"]


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def scan_page(label, url):
    r = try_get(label, url, allow_redirects=True)
    if r is None or r.status_code != 200:
        return
    text = r.text
    hits = [k for k in SECTOR_KEYWORDS if k in text.lower()]
    print(f"  sector keyword hits: {hits}", file=sys.stderr)
    links = sorted(set(re.findall(r'href="([^"]+)"', text)))
    data_links = [l for l in links if any(ext in l.lower() for ext in (".csv", ".xlsx", ".json", "/api/", "api."))][:20]
    print(f"  data-looking links: {data_links}", file=sys.stderr)


CANDIDATES = [
    ("Denmark - Energi Data Service API", "https://api.energidataservice.dk/dataset"),
    ("Denmark - Energinet transparency", "https://energinet.dk/gas/gasdata"),
    ("Romania - Transgaz transparency", "https://www.transgaz.ro/en/transparency"),
    ("Portugal - REN gas data", "https://www.ren.pt/en-us/what-we-do/natural-gas/system-data"),
    ("Belgium - Fluxys transparency", "https://www.fluxys.com/en/transparency"),
    ("Poland - GAZ-SYSTEM transparency", "https://www.gaz-system.pl/en/customers/transparency-platform.html"),
    ("Italy - Snam transparency", "https://www.snam.it/en/transport/transparency/"),
    ("Netherlands - GTS transparency", "https://www.gasunietransportservices.nl/en/transparency"),
]

if __name__ == "__main__":
    for label, url in CANDIDATES:
        try:
            scan_page(label, url)
        except Exception as e:
            print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
