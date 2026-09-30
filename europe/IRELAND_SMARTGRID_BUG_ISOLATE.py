"""
ireland_smartgrid.py's first live run returned 0 rows for EVERY single
chunk (47 chunks x 3 chart types, from the very first request) despite
IRELAND_SMARTGRID_RANGE_INSPECT.py confirming a wide (31-day) date span
works minutes earlier. The one difference never actually tested: that
inspect script's wide-range test used a single area ("demandactual"),
while the production script always joins multiple areas with a comma
("demandactual,demandforecast"). Isolating exactly which combination
breaks:
  1. wide range + single area (confirmed working already, sanity check)
  2. wide range + comma-joined areas (never tested - main suspect)
  3. same as #2 but using a requests.Session() with a small delay
     between calls (in case it's actually a rapid-fire/bot-detection
     issue, not a query-format issue)
"""
import sys
import time

import requests

BASE = "https://www.smartgriddashboard.com/api/chart/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://www.smartgriddashboard.com/",
}
TIMEOUT = (10, 45)


def try_get(label, session, params):
    print(f"\n{'=' * 70}\n{label}: {params}\n{'=' * 70}", file=sys.stderr)
    getter = session.get if session is not None else requests.get
    r = getter(BASE, headers=HEADERS, params=params, timeout=TIMEOUT)
    print(f"  status={r.status_code}, final url={r.url}", file=sys.stderr)
    data = r.json()
    rows = data.get("Rows", [])
    print(f"  {len(rows)} rows", file=sys.stderr)
    if rows:
        print(f"  first: {rows[0]}", file=sys.stderr)


if __name__ == "__main__":
    try_get("1. wide range, single area (sanity check - known working)", None, {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Jan-2020", "areas": "demandactual",
    })
    try_get("2. wide range, COMMA-JOINED areas (main suspect)", None, {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Jan-2020", "areas": "demandactual,demandforecast",
    })
    session = requests.Session()
    try_get("3. same as #2, with a Session", session, {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "01-Jan-2020", "dateTo": "31-Jan-2020", "areas": "demandactual,demandforecast",
    })
    time.sleep(1)
    try_get("4. same as #2, single-day span this time (known working shape)", None, {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "15-Jan-2020", "dateTo": "15-Jan-2020", "areas": "demandactual,demandforecast",
    })
    # Exact reproduction of production's very first failing chunk.
    try_get("5. exact repro of the first failing production chunk", None, {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "15-Jan-2015", "dateTo": "14-Apr-2015", "areas": "demandactual,demandforecast",
    })
