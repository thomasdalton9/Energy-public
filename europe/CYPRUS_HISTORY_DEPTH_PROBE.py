"""
Binary-search probe for how far back TSOC's historical archive
(archive-total-daily-system-generation-on-the-transmission-system)
actually has real data, as opposed to how far back the ARCHIVE WEB PAGE
itself has existed (published 2018-04-25, per its own metadata - not
necessarily when the underlying data starts).

Reuses cyprus_generation_mix.py's fetch_day() directly rather than
reimplementing the request/parse logic, so this probe is testing the
exact same code path the real script uses.

Strategy: confirm a known-good recent date and a plausible known-bad
old date, then binary search the boundary between them. Also spot-checks
a few fixed calendar years along the way for a sanity picture, since a
single boundary date doesn't reveal whether there are GAPS in the
middle of the range (e.g. an outage, a site migration) - binary search
alone assumes monotonic availability, which might not hold.
"""

import sys
import time
from datetime import date, timedelta

sys.path.insert(0, "europe")
import cyprus_generation_mix as cgm
import requests

REQUEST_PAUSE_SECONDS = 0.3


def check_date(session, day):
    try:
        rows, _ = cgm.fetch_day(session, day)
        return len(rows) > 0, len(rows)
    except requests.RequestException as e:
        return None, str(e)
    finally:
        time.sleep(REQUEST_PAUSE_SECONDS)


def main():
    session = requests.Session()

    print("=== Sanity spot-checks across candidate years ===")
    for year in (2010, 2012, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025):
        day = date(year, 6, 1)
        ok, detail = check_date(session, day)
        print(f"  {day.isoformat()}: {'OK (' + str(detail) + ' rows)' if ok else 'NO DATA (' + str(detail) + ')' if ok is False else 'ERROR: ' + str(detail)}")

    print("\n=== Binary search for the earliest date with real data ===")
    # lower = a date we expect to have NO data; upper = a date we know HAS data
    lower = date(2005, 1, 1)
    upper = date(2025, 6, 1)

    lower_ok, _ = check_date(session, lower)
    upper_ok, _ = check_date(session, upper)
    print(f"  lower bound {lower.isoformat()}: {'HAS DATA (unexpected!)' if lower_ok else 'no data (as expected)'}")
    print(f"  upper bound {upper.isoformat()}: {'has data (as expected)' if upper_ok else 'NO DATA (unexpected!)'}")

    if lower_ok:
        print("  lower bound unexpectedly has data - binary search assumption invalid, stopping here")
        return
    if not upper_ok:
        print("  upper bound unexpectedly has no data - binary search assumption invalid, stopping here")
        return

    while (upper - lower).days > 1:
        mid = lower + (upper - lower) // 2
        ok, detail = check_date(session, mid)
        print(f"  probing {mid.isoformat()}: {'HAS DATA' if ok else 'no data'} ({detail})")
        if ok:
            upper = mid
        else:
            lower = mid

    print(f"\nEarliest date with data appears to be around: {upper.isoformat()}")
    print(f"(last confirmed no-data date: {lower.isoformat()})")

    print("\n=== Confirming a few days right around that boundary ===")
    for offset in (-3, -2, -1, 0, 1, 2, 3):
        day = upper + timedelta(days=offset)
        ok, detail = check_date(session, day)
        print(f"  {day.isoformat()}: {'HAS DATA (' + str(detail) + ' rows)' if ok else 'no data'}")


if __name__ == "__main__":
    main()
