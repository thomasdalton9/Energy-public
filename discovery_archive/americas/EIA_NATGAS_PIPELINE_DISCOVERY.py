"""
One-off discovery probe for two things ahead of building anything:

1. Does EIA's structured API (not just its narrative reports) expose a
   Waha (Permian-area) natural gas spot price, or only Henry Hub? Same
   question as the coal discovery just answered for coal prices -
   EIA's *narrative* reports show many hub prices, but that doesn't
   mean they're in the queryable API as clean series.

2. Does EIA have any structured series for natural gas pipeline
   capacity or flows (e.g. Permian takeaway capacity) - as opposed to
   only being covered in narrative reports/PDFs (which is what turned
   out to be true for coal's weekly benchmark prices).

Lists natural-gas's own sub-routes first (same parent-route trick used
for coal/electricity earlier), then drills into whichever look
price/pipeline/capacity-related.

Needs EIA_API_KEY in the environment (same repo secret as the other
EIA discovery probes).
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

    print("=== 1. Top-level route list: natural-gas ===")
    top = get("natural-gas")
    routes = top.get("response", {}).get("routes", [])
    print(json.dumps(routes, indent=2))

    for r in routes:
        route_id = r["id"]
        print(f"\n=== 2. Sub-routes: natural-gas/{route_id} ===")
        try:
            meta = get(f"natural-gas/{route_id}")
            sub = meta.get("response", {}).get("routes")
            if sub is not None:
                print(json.dumps(sub, indent=2))
            else:
                print(json.dumps(meta, indent=2)[:1500])
        except requests.RequestException as e:
            print(f"  failed: {e}")

    print("\n=== 3. Looking specifically at natural-gas/pri (price) sub-routes and facets ===")
    try:
        pri = get("natural-gas/pri")
        print(json.dumps(pri, indent=2)[:2000])
        pri_routes = pri.get("response", {}).get("routes", [])
        for r in pri_routes:
            rid = r["id"]
            print(f"\n--- natural-gas/pri/{rid} metadata ---")
            meta = get(f"natural-gas/pri/{rid}")
            print(json.dumps(meta, indent=2)[:2000])
    except requests.RequestException as e:
        print(f"  failed: {e}")

    print("\n=== 4. Searching all natural-gas sub-routes for 'pipe'/'capacity'/'flow' in id or name ===")
    keywords = ["pipe", "capacit", "flow", "transport"]
    for r in routes:
        route_id = r["id"]
        name = r.get("name", "")
        if any(kw in route_id.lower() or kw in name.lower() for kw in keywords):
            print(f"  MATCH: {r}")


if __name__ == "__main__":
    main()
