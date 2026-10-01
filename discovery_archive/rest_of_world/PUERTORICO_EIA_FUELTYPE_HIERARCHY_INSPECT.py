"""
Follow-up to PUERTORICO_EIA_SECTOR_NAMES_INSPECT.py (which confirmed a
clean, mutually-exclusive sector breakdown: sectors 1-7 sum exactly to
sector 99 "All Sectors"). Fuel types are murkier - PUERTORICO_EIA_
INSPECT.py's full fuel-type list mixes what look like aggregates (ALL,
AOR, COW, FOS, NGO, PEL, PET, REN) with what look like leaf categories
(COL, BIS, BIT, DFO, HYC, LFG, MLG, OOG, OTH, RFO, SPV, SUN, WAS, WND,
WNT, WOO, BIO) - some "leaf" candidates might themselves be sums of
others (e.g. COL vs BIT/BIS, SUN vs SPV, WND vs WNT), which would
double-count a naive sum.

Fetching every fueltypeid for sectorid=99 (All Sectors) for one recent
period and checking the arithmetic directly: which subset of
non-aggregate-looking ids actually sums to the "ALL" total, without
overlap.
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
        "facets[sectorid][]": "99",
        "sort[0][column]": "period",
        "sort[0][direction]": "desc",
        "length": 50,
    }
    r = requests.get(URL, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    rows = r.json()["response"]["data"]
    if not rows:
        print("no rows", file=sys.stderr)
        return
    latest_period = rows[0]["period"]
    latest_rows = [row for row in rows if row["period"] == latest_period]
    print(f"period {latest_period}, sectorid=99 (All Sectors), {len(latest_rows)} fuel type rows:",
          file=sys.stderr)
    values = {}
    for row in sorted(latest_rows, key=lambda r: r["fueltypeid"]):
        gen = row["generation"]
        values[row["fueltypeid"]] = float(gen) if gen not in (None, "") else None
        print(f"  {row['fueltypeid']:5s} ({row['fuelTypeDescription']:40s}): {gen}", file=sys.stderr)

    total = values.get("ALL")
    print(f"\nALL = {total}", file=sys.stderr)

    known_aggregates = {"ALL", "AOR", "COW", "FOS", "NGO", "PEL", "PET", "REN"}
    candidate_leaves = {k: v for k, v in values.items() if k not in known_aggregates and v is not None}
    leaf_sum = sum(candidate_leaves.values())
    print(f"sum of non-aggregate-looking ids {sorted(candidate_leaves)}: {leaf_sum}", file=sys.stderr)
    print(f"matches ALL: {abs(leaf_sum - total) < 0.01 if total is not None else 'N/A (ALL missing)'}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
