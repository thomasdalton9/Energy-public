"""
Follow-up to SRILANKA_PUCSL_API_PARAMS_INSPECT.py: every GET attempt on
actual-system-dispatch, reservoir/storage-rainfall, daily-generation-
statistics, and bulk-supply-tariff/* returned HTTP 500 regardless of
query params tried. The frontend SPA might call these via POST with a
JSON body instead - trying that before giving up on them.
"""
import json
import sys

import requests

BASE = "https://gendata.pucsl.gov.lk/api"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Referer": "https://gendata.pucsl.gov.lk/home",
}
TIMEOUT = (10, 45)

ATTEMPTS = [
    ("actual-system-dispatch", {"year": 2023, "month": 6, "day": 1}),
    ("actual-system-dispatch", {"date": "2023-06-01", "type": "daily"}),
    ("reservoir/storage-rainfall", {"year": 2023, "month": 6}),
    ("daily-generation-statistics", {"year": 2023, "month": 6}),
    ("bulk-supply-tariff/capacity", {"year": 2023}),
    ("bulk-supply-tariff/energy", {"year": 2023}),
]


def inspect_post(path, body):
    url = f"{BASE}/{path}"
    print(f"\n{'=' * 70}\nPOST {url}  body={body}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.post(url, headers=HEADERS, json=body, timeout=TIMEOUT)
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
    for path, body in ATTEMPTS:
        inspect_post(path, body)
