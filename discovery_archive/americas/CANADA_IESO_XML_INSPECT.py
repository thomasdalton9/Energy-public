"""
Follow-up to CANADA_IESO_DISCOVERY2.py, which confirmed reports-public.
ieso.ca is alive with real, current Generator Output and Capability
XML reports (PUB_GenOutputCapability.xml, updated today). Fetching the
current one and printing its real structure before building a
production script.
"""
import sys

import requests

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 45)

URL = "https://reports-public.ieso.ca/public/GenOutputCapability/PUB_GenOutputCapability.xml"

r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
text = r.text
print(f"\nfirst 3000 chars:\n{text[:3000]}", file=sys.stderr)
print(f"\nlast 1000 chars:\n{text[-1000:]}", file=sys.stderr)
