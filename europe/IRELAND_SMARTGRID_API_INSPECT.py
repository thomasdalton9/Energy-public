"""
Follow-up to IRELAND_TURKEY_POWER_PLAYWRIGHT_DISCOVERY.py: Playwright
network capture of smartgriddashboard.com (EirGrid/SONI's all-island
Smart Grid Dashboard) found its real underlying API:

    https://www.smartgriddashboard.com/api/chart/?region=ALL&chartType=demand
        &dateRange=day&dateFrom=30-Sep-2026&dateTo=30-Sep-2026&areas=demandactual,demandforecast

No login, no key. The page's own nav lists chart types beyond Demand:
Generation, Wind Generation, Solar Generation, Interconnection,
Frequency, Imbalance Price/Volume, SNSP, CO2 - checking whether
chartType=generation (plus guesses for a few others) works the same
way, what its "areas" values/fuel categories are, and how far back
dateFrom can go (today's date only proves "today" works).
"""
import json
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
        print(f"  body: {r.text[:500]}", file=sys.stderr)
        return
    try:
        data = r.json()
    except ValueError:
        print(f"  not JSON: {r.text[:500]}", file=sys.stderr)
        return
    if isinstance(data, dict):
        print(f"  dict keys: {list(data.keys())}", file=sys.stderr)
        rows = data.get("Rows") or data.get("rows") or data.get("data")
        if isinstance(rows, list):
            print(f"  {len(rows)} rows, first: {json.dumps(rows[0])[:400]}", file=sys.stderr)
            print(f"  last: {json.dumps(rows[-1])[:400]}", file=sys.stderr)
        else:
            print(f"  preview: {json.dumps(data)[:800]}", file=sys.stderr)
    elif isinstance(data, list):
        print(f"  list of {len(data)}, first: {json.dumps(data[0])[:400]}", file=sys.stderr)
        print(f"  last: {json.dumps(data[-1])[:400]}", file=sys.stderr)


if __name__ == "__main__":
    # Today, known-working chartType (sanity check).
    try_get("demand (today, known-working)", {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "30-Sep-2026", "dateTo": "30-Sep-2026", "areas": "demandactual,demandforecast",
    })
    # Fuel-mix generation guesses.
    for areas in ["generationactual", "windactual,generationactual", "fuelmix"]:
        try_get(f"generation, areas={areas}", {
            "region": "ALL", "chartType": "generation", "dateRange": "day",
            "dateFrom": "30-Sep-2026", "dateTo": "30-Sep-2026", "areas": areas,
        })
    try_get("wind", {
        "region": "ALL", "chartType": "wind", "dateRange": "day",
        "dateFrom": "30-Sep-2026", "dateTo": "30-Sep-2026", "areas": "windactual,windforecast",
    })
    try_get("co2", {
        "region": "ALL", "chartType": "co2", "dateRange": "day",
        "dateFrom": "30-Sep-2026", "dateTo": "30-Sep-2026", "areas": "co2intensity",
    })
    # Historical depth check - a date years in the past.
    try_get("demand, historical date (2020-01-15)", {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "15-Jan-2020", "dateTo": "15-Jan-2020", "areas": "demandactual,demandforecast",
    })
    try_get("demand, historical date (2015-01-15)", {
        "region": "ALL", "chartType": "demand", "dateRange": "day",
        "dateFrom": "15-Jan-2015", "dateTo": "15-Jan-2015", "areas": "demandactual,demandforecast",
    })
