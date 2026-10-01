"""
Follow-up to SRILANKA_PUCSL_API_INSPECT.py: actual-system-dispatch,
reservoir/storage-rainfall, daily-generation-statistics, and both
bulk-supply-tariff endpoints all returned HTTP 500 with no params,
while metadata/latest-year confirmed the site's own "latest year" is
2023 - guessing these need a date/year query param the frontend
normally supplies. Trying common param names/values before giving up
on them.
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

ATTEMPTS = [
    ("actual-system-dispatch", {"date": "2023-06-01"}),
    ("actual-system-dispatch", {"year": 2023, "month": 6}),
    ("reservoir/storage-rainfall", {"year": 2023}),
    ("reservoir/storage-rainfall", {"date": "2023-06-01"}),
    ("daily-generation-statistics", {"year": 2023}),
    ("daily-generation-statistics", {"date": "2023-06-01"}),
    ("daily-generation-statistics", {"from": "2023-01-01", "to": "2023-01-31"}),
    ("bulk-supply-tariff/capacity", {"year": 2023}),
    ("bulk-supply-tariff/energy", {"year": 2023}),
]


def inspect(path, params):
    url = f"{BASE}/{path}"
    print(f"\n{'=' * 70}\n{url}  params={params}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, params=params, timeout=TIMEOUT)
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return
    print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
    if r.status_code != 200:
        print(f"  body: {r.text[:300]}", file=sys.stderr)
        return
    try:
        data = r.json()
    except ValueError:
        print(f"  not JSON: {r.text[:300]}", file=sys.stderr)
        return
    print(f"  SUCCESS - preview: {json.dumps(data, indent=2)[:1200]}", file=sys.stderr)


if __name__ == "__main__":
    for path, params in ATTEMPTS:
        inspect(path, params)
