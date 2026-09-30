"""
One-off discovery probe for EIA-930 (Hourly Electric Grid Monitor) via
the EIA API v2, ahead of building EIA930_FUEL_MIX_DAILY.py.

Unlike MISO/FRED, this needs a free api.eia.gov key - read from the
EIA_API_KEY environment variable (a GitHub Actions secret in this
repo's workflow, never committed or logged). This script never prints
the key itself.

Checks, in order:
  1. Metadata for electricity/rto/daily-fuel-type-data - confirms the
     route exists and lists valid facets (fueltype codes, respondent
     codes including the "US48" Lower-48 total).
  2. A small real data pull: US48, last 10 days, all fuel types - to
     see the actual row shape (columns, units, one row per
     respondent/fueltype/period).
  3. Same for electricity/rto/daily-region-data (demand / net
     generation / interchange totals, not split by fuel type) as a
     second candidate dataset.
"""

import json
import os
import sys

import requests

API_KEY = os.environ.get("EIA_API_KEY")
BASE = "https://api.eia.gov/v2"
TIMEOUT = (10, 30)


def get(path, params):
    params = dict(params)
    params["api_key"] = API_KEY
    r = requests.get(f"{BASE}/{path}", params=params, timeout=TIMEOUT)
    print(f"GET {path} -> {r.status_code}", file=sys.stderr)
    r.raise_for_status()
    return r.json()


def main():
    if not API_KEY:
        print("EIA_API_KEY is not set in the environment - aborting.", file=sys.stderr)
        sys.exit(1)
    print(f"EIA_API_KEY present, length {len(API_KEY)} (not printing the value)", file=sys.stderr)

    print("\n=== 1. Metadata: electricity/rto/daily-fuel-type-data ===")
    meta = get("electricity/rto/daily-fuel-type-data", {})
    print(json.dumps(meta, indent=2)[:3000])

    print("\n=== 2. Data pull: daily-fuel-type-data, respondent=US48, last 10 rows ===")
    data = get(
        "electricity/rto/daily-fuel-type-data/data",
        {
            "frequency": "daily",
            "data[0]": "value",
            "facets[respondent][]": "US48",
            "sort[0][column]": "period",
            "sort[0][direction]": "desc",
            "length": 10,
        },
    )
    print(json.dumps(data, indent=2)[:5000])

    print("\n=== 3. Metadata: electricity/rto/daily-region-data ===")
    meta2 = get("electricity/rto/daily-region-data", {})
    print(json.dumps(meta2, indent=2)[:3000])

    print("\n=== 4. Data pull: daily-region-data, respondent=US48, last 10 rows ===")
    data2 = get(
        "electricity/rto/daily-region-data/data",
        {
            "frequency": "daily",
            "data[0]": "value",
            "facets[respondent][]": "US48",
            "sort[0][column]": "period",
            "sort[0][direction]": "desc",
            "length": 10,
        },
    )
    print(json.dumps(data2, indent=2)[:5000])


if __name__ == "__main__":
    main()
