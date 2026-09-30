"""
Follow-up to IRELAND_SMARTGRID_API_INSPECT.py: chartType=demand
confirmed working back to at least 2015-01-15 for a SINGLE day
(dateFrom == dateTo). Checking whether dateFrom/dateTo can span a
wider range in one call (dateRange=week/month/year, or just
dateFrom != dateTo with dateRange=day) - if so, a historical backfill
needs far fewer requests than Mexico's day-by-day CENACE loop.
"""
import sys

import requests

BASE = "https://www.smartgriddashboard.com/api/chart/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://www.smartgriddashboard.com/",
}
TIMEOUT = (10, 45)


def try_get(label, params):
    print(f"\n{'=' * 70}\n{label}: {params}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(BASE, headers=HEADERS, params=params, timeout=TIMEOUT)
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return
    print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
    if r.status_code != 200:
        print(f"  body: {r.text[:300]}", file=sys.stderr)
        return
    data = r.json()
    rows = data.get("Rows", [])
    print(f"  {len(rows)} rows", file=sys.stderr)
    if rows:
        print(f"  first: {rows[0]}", file=sys.stderr)
        print(f"  last: {rows[-1]}", file=sys.stderr)
        dates = sorted(set(r["EffectiveTime"][:11] for r in rows))
        print(f"  distinct dates present: {len(dates)} ({dates[:3]}...{dates[-3:]})", file=sys.stderr)


if __name__ == "__main__":
    try_get("dateRange=day, dateFrom != dateTo (1 month span)", {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Jan-2020", "areas": "demandactual",
    })
    try_get("dateRange=week", {
        "region": "ALL", "chartType": "demand", "dateRange": "week",
        "dateFrom": "01-Jan-2020", "dateTo": "07-Jan-2020", "areas": "demandactual",
    })
    try_get("dateRange=month", {
        "region": "ALL", "chartType": "demand", "dateRange": "month",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Jan-2020", "areas": "demandactual",
    })
    try_get("dateRange=year", {
        "region": "ALL", "chartType": "demand", "dateRange": "year",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Dec-2020", "areas": "demandactual",
    })
    try_get("generation, all known FieldName guesses as separate areas", {
        "region": "ALL", "chartType": "fuel_mix", "dateRange": "day",
        "dateFrom": "30-Sep-2026", "dateTo": "30-Sep-2026", "areas": "fuelmix",
    })
