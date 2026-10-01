"""
Follow-up to PUERTORICO_EIA_INSPECT.py: that script's sector summary had
a bug (grouped on "sectorName", a key that doesn't exist - the real key
is "sectorDescription"), so it printed None for every sector's label.
Re-fetching one period's rows (fueltypeid=ALL only, to cut volume) to
get the real sectorid -> sectorDescription mapping before building the
production script.
"""
import os
import sys

import requests

API_KEY = os.environ.get("EIA_API_KEY")
URL = "https://api.eia.gov/v2/electricity/electric-power-operational-data/data/"
TIMEOUT = (10, 45)


def main():
    if not API_KEY:
        print("EIA_API_KEY not set - aborting", file=sys.stderr)
        sys.exit(1)
    params = {
        "api_key": API_KEY,
        "frequency": "monthly",
        "data[]": "generation",
        "facets[location][]": "PR",
        "facets[fueltypeid][]": "ALL",
        "sort[0][column]": "period",
        "sort[0][direction]": "desc",
        "length": 50,
    }
    r = requests.get(URL, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    rows = r.json()["response"]["data"]
    mapping = sorted(set((row["sectorid"], row["sectorDescription"]) for row in rows))
    print(f"{len(rows)} rows, sectorid -> sectorDescription mapping (fueltypeid=ALL only):", file=sys.stderr)
    for sid, desc in mapping:
        print(f"  {sid}: {desc}", file=sys.stderr)
    print("\nmost recent period's rows:", file=sys.stderr)
    latest_period = rows[0]["period"]
    for row in rows:
        if row["period"] == latest_period:
            print(f"  {row}", file=sys.stderr)


if __name__ == "__main__":
    main()
