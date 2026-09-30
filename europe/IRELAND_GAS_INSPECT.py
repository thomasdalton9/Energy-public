"""
Inspect Gas Networks Ireland's Daily Gas Demand CSV (URL confirmed live
via data.gov.ie's CKAN API in IRELAND_TURKEY_DISCOVERY.py) to see its
real column names/structure before building the production script.
"""
import sys

import requests

URL = "https://www.gasnetworks.ie/sites/default/files/2026-08/Gas_Demand.csv"
HEADERS = {"User-Agent": "Mozilla/5.0"}

r = requests.get(URL, headers=HEADERS, timeout=(10, 45))
print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
text = r.text
lines = text.splitlines()
print(f"{len(lines)} lines total", file=sys.stderr)
print("\n--- first 10 lines ---", file=sys.stderr)
for line in lines[:10]:
    print(line, file=sys.stderr)
print("\n--- last 10 lines ---", file=sys.stderr)
for line in lines[-10:]:
    print(line, file=sys.stderr)
