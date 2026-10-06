"""
Follow-up to EIA_NATGAS_PIPELINE_DISCOVERY.py: that probe found
natural-gas/pri/fut has daily/weekly frequency and a "duoarea" facet -
promising, since Henry Hub alone wouldn't need a location facet. This
checks the actual duoarea facet values for a Waha/Permian entry, and
pulls a sample of real data if one exists.
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
        print("EIA_API_KEY is not set - aborting.", file=sys.stderr)
        sys.exit(1)

    for route in ("natural-gas/pri/spt", "natural-gas/pri/sum"):
        try:
            fac = get(f"{route}/facet/duoarea").get("response", {}).get("facets", [])
            print(f"=== duoarea {route}: {len(fac)} ===")
            print(json.dumps(fac, indent=2)[:2500])
        except Exception as e:
            print(route, "ERR", e)
    print("=== duoarea facet values: natural-gas/pri/fut ===")
    facet = get("natural-gas/pri/fut/facet/duoarea")
    values = facet.get("response", {}).get("facets", [])
    print(f"total: {len(values)}")
    print(json.dumps(values, indent=2))

    waha_like = [v for v in values if "waha" in v.get("name", "").lower() or "waha" in v.get("id", "").lower()
                 or "permian" in v.get("name", "").lower()]
    print(f"\nWaha/Permian matches: {waha_like}")

    print("\n=== series facet values: natural-gas/pri/fut ===")
    series_facet = get("natural-gas/pri/fut/facet/series")
    print(json.dumps(series_facet.get("response", {}).get("facets", []), indent=2)[:3000])

    if waha_like:
        target = waha_like[0]["id"]
        print(f"\n=== sample data pull: duoarea={target} ===")
        data = get("natural-gas/pri/fut/data", {
            "frequency": "daily",
            "data[0]": "value",
            "facets[duoarea][]": target,
            "sort[0][column]": "period",
            "sort[0][direction]": "desc",
            "length": 10,
        })
        print(json.dumps(data, indent=2)[:3000])


if __name__ == "__main__":
    main()
