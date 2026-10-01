"""
Follow-up to GATUN_RHINE_PANAMA_DISCOVERY.py's puerto_rico_eia() check,
which confirmed electricity/electric-power-operational-data works for
facets[location][]=PR but only sampled 20 rows unfiltered - those all
happened to be the "Electric Utility" sector. Widening the query here:
paginating through every row for PR (no sector filter) to see the real
full set of sectors (e.g. independent power producers) and fuel types
reported, and the actual date range available, before building the
production script.
"""
import os
import sys

import requests

API_KEY = os.environ.get("EIA_API_KEY")
URL = "https://api.eia.gov/v2/electricity/electric-power-operational-data/data/"
TIMEOUT = (10, 45)


def fetch_page(offset, length=5000):
    params = {
        "api_key": API_KEY,
        "frequency": "monthly",
        "data[]": "generation",
        "facets[location][]": "PR",
        "sort[0][column]": "period",
        "sort[0][direction]": "asc",
        "offset": offset,
        "length": length,
    }
    r = requests.get(URL, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def main():
    if not API_KEY:
        print("EIA_API_KEY not set - aborting", file=sys.stderr)
        sys.exit(1)

    rows = []
    offset = 0
    while True:
        payload = fetch_page(offset)
        page_rows = payload["response"]["data"]
        rows.extend(page_rows)
        total = int(payload["response"]["total"])
        offset += len(page_rows)
        print(f"  fetched {offset}/{total} rows...", file=sys.stderr)
        if not page_rows or offset >= total:
            break

    print(f"\n{len(rows)} total rows for PR", file=sys.stderr)
    if not rows:
        return

    sectors = sorted(set((r.get("sectorid"), r.get("sectorName")) for r in rows))
    print(f"\nsectors present: {sectors}", file=sys.stderr)

    fueltypes = sorted(set((r.get("fueltypeid"), r.get("fuelTypeDescription")) for r in rows))
    print(f"\nfuel types present: {fueltypes}", file=sys.stderr)

    periods = sorted(set(r.get("period") for r in rows))
    print(f"\nperiod range: {periods[0]} to {periods[-1]} ({len(periods)} distinct periods)", file=sys.stderr)

    print(f"\nfirst 3 rows raw: {rows[:3]}", file=sys.stderr)
    print(f"\nsample row keys: {list(rows[0].keys())}", file=sys.stderr)


if __name__ == "__main__":
    main()
