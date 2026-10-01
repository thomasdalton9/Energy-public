"""
Follow-up to SRILANKA_PUCSL_PLAYWRIGHT_DISCOVERY.py, which used
Playwright network capture (same technique as Bolivia's CNDC) to find
gendata.pucsl.gov.lk's real underlying REST API. Confirmed endpoints
found that way (no key needed):

    gendata.pucsl.gov.lk/api/actual-system-dispatch
    gendata.pucsl.gov.lk/api/reservoir/storage-rainfall
    gendata.pucsl.gov.lk/api/daily-generation-statistics
    gendata.pucsl.gov.lk/api/metadata/power-plants
    gendata.pucsl.gov.lk/api/metadata/power-plant-complexes
    gendata.pucsl.gov.lk/api/bulk-supply-tariff/capacity
    gendata.pucsl.gov.lk/api/bulk-supply-tariff/energy
    gendata.pucsl.gov.lk/api/metadata/latest-year

Hitting each directly with plain requests (no browser needed) to see
its real JSON shape/fields and whether it needs any query params
(date range, year) before building the production script.
"""
import json
import sys

import requests

BASE = "https://gendata.pucsl.gov.lk/api"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://gendata.pucsl.gov.lk/home",
}
TIMEOUT = (10, 45)

ENDPOINTS = [
    "actual-system-dispatch",
    "reservoir/storage-rainfall",
    "daily-generation-statistics",
    "metadata/power-plants",
    "metadata/power-plant-complexes",
    "bulk-supply-tariff/capacity",
    "bulk-supply-tariff/energy",
    "metadata/latest-year",
]


def inspect(path):
    url = f"{BASE}/{path}"
    print(f"\n{'=' * 70}\n{url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return
    print(f"  status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
          file=sys.stderr)
    if r.status_code != 200:
        print(f"  body: {r.text[:500]}", file=sys.stderr)
        return
    try:
        data = r.json()
    except ValueError:
        print(f"  not JSON, body: {r.text[:500]}", file=sys.stderr)
        return
    if isinstance(data, list):
        print(f"  list of {len(data)} items", file=sys.stderr)
        if data:
            print(f"  first item: {json.dumps(data[0], indent=2)[:1500]}", file=sys.stderr)
            print(f"  last item: {json.dumps(data[-1], indent=2)[:1500]}", file=sys.stderr)
    elif isinstance(data, dict):
        print(f"  dict keys: {list(data.keys())}", file=sys.stderr)
        print(f"  preview: {json.dumps(data, indent=2)[:2000]}", file=sys.stderr)
    else:
        print(f"  value: {data!r}", file=sys.stderr)


if __name__ == "__main__":
    for path in ENDPOINTS:
        inspect(path)
