"""
One-off discovery probe for EIA coal price data via the EIA API v2,
ahead of building a weekly coal-basin-price puller.

Not yet confirmed: whether the weekly "Coal Markets" benchmark spot
price (Central Appalachia, Northern Appalachia, Illinois Basin, Powder
River Basin, Uinta Basin - $/ton) is exposed as a structured API v2
series at all (like EIA-930/STEO) or only published as a report/webpage
(like the Drilling Productivity Report, which turned out to be a plain
file download instead - see EIA_DPR_DISCOVERY.py). This probe first
lists coal's own sub-routes (querying a parent route with no further
path returns its child routes, same trick used for electricity/rto
earlier), then drills into whichever looks like price data.

Needs EIA_API_KEY in the environment (same repo secret as
EIA930_DISCOVERY.py).
"""

import json
import os
import sys

import requests

API_KEY = os.environ.get("EIA_API_KEY")
BASE = "https://api.eia.gov/v2"
TIMEOUT = (10, 30)


def get(path, params=None):
    params = dict(params or {})
    params["api_key"] = API_KEY
    r = requests.get(f"{BASE}/{path}", params=params, timeout=TIMEOUT)
    print(f"GET {path} -> {r.status_code}", file=sys.stderr)
    r.raise_for_status()
    return r.json()


def main():
    if not API_KEY:
        print("EIA_API_KEY is not set in the environment - aborting.", file=sys.stderr)
        sys.exit(1)

    print("=== 1. Top-level route list: coal ===")
    top = get("coal")
    print(json.dumps(top, indent=2))

    routes = top.get("response", {}).get("routes", [])
    price_like = [r for r in routes if "price" in r.get("id", "").lower() or "price" in r.get("name", "").lower()]
    print(f"\nprice-like sub-routes: {price_like}")

    for r in routes:
        route_id = r["id"]
        print(f"\n=== 2. Metadata: coal/{route_id} ===")
        try:
            meta = get(f"coal/{route_id}")
            print(json.dumps(meta, indent=2)[:2500])
        except requests.RequestException as e:
            print(f"  failed: {e}")

    if price_like:
        target = price_like[0]["id"]
        print(f"\n=== 3. Sample data pull: coal/{target}/data ===")
        try:
            data = get(f"coal/{target}/data", {"length": 10, "sort[0][column]": "period", "sort[0][direction]": "desc"})
            print(json.dumps(data, indent=2)[:4000])
        except requests.RequestException as e:
            print(f"  failed: {e}")


if __name__ == "__main__":
    main()
