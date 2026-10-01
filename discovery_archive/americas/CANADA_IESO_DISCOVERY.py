"""
Discovery for Canada's power generation data. IESO (Independent
Electricity System Operator, Ontario) is the strongest lead from
general knowledge - it publishes real open data reports, historically
via CSV/XML at reports.ieso.ca, plus a newer API. Checking known report
URLs (generator output by fuel type) and the newer data portal.

Ontario is one province, not all of Canada - checking if IESO's output
is presented as "Ontario" explicitly (it is, by design) as a scoping
note for whatever gets built.
"""
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 45)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
              file=sys.stderr)
        if r.status_code == 200:
            print(f"  first 500 chars: {r.text[:500]}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


if __name__ == "__main__":
    # Classic IESO public reports (no key) - generator output/capability by fuel type.
    try_get("IESO GenOutputCapability today", "http://reports.ieso.ca/public/GenOutputCapability/")
    try_get("IESO GenOutputbyFuelHourly today",
            "http://reports.ieso.ca/public/GenOutputbyFuelHourly/")
    try_get("IESO data portal", "https://www.ieso.ca/en/power-data")
    try_get("IESO newer API root guess", "https://www.ieso.ca/api/")
