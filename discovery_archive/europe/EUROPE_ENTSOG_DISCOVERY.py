"""
ENTSOG Transparency Platform probe for the Europe dashboard (no API key needed).

Goal: find out how to get country-level gas flows by category (distribution /
final consumers / power plants / storage / LNG / production / cross-border)
from the platform, so europe/EUROPE_ENTSOG_GAS.py can chart gas demand per
country. Prints endpoint status, field names and sample rows - read the log.

Usage: python3 EUROPE_ENTSOG_DISCOVERY.py
"""
import json
import sys

import requests

BASE = "https://transparency.entsog.eu/api/v1/"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}


def show(label, endpoint, **params):
    print(f"\n=== {label}: {endpoint} {params}", flush=True)
    try:
        r = requests.get(BASE + endpoint, headers=H, params=params, timeout=(15, 120))
    except requests.RequestException as e:
        print("  ERROR", type(e).__name__, e)
        return None
    print("  status", r.status_code, "bytes", len(r.content))
    if r.status_code != 200:
        print("  body:", r.text[:300])
        return None
    try:
        j = r.json()
    except ValueError:
        print("  not JSON:", r.text[:300])
        return None
    for k, v in j.items():
        if isinstance(v, list):
            print(f"  list '{k}': {len(v)} rows")
            for row in v[:3]:
                print("   ", json.dumps(row)[:700])
            return v
        print("  ", k, str(v)[:200])
    return None


show("operators", "operators", limit=3)
show("balancing zones", "balancingzones", limit=3)
show("points", "interconnections", limit=3)
show("point types", "Interconnections", limit=2)
rows = show("aggregated (month, DE)", "AggregatedData", indicator="Physical Flow", periodType="month",
            **{"from": "2025-01-01", "to": "2025-03-01", "countryKey": "DE", "limit": 20})
show("aggregated (month, all)", "AggregatedData", indicator="Physical Flow", periodType="month",
     **{"from": "2025-01-01", "to": "2025-02-01", "limit": 10})
show("aggregated lowercase", "aggregatedData", indicator="Physical Flow", periodType="month",
     **{"from": "2025-01-01", "to": "2025-02-01", "limit": 10})
show("operational by tsoCountry", "operationalData", indicator="Physical Flow", periodType="month",
     **{"from": "2025-01-01", "to": "2025-02-01", "countryKey": "DE", "limit": 10})
show("operational point category", "operationalData", indicator="Physical Flow", periodType="day",
     **{"from": "2025-01-01", "to": "2025-01-03", "limit": 10, "pointKey": "ITP-00495"})
show("operational with adjacentSystem filter", "operationalData", indicator="Physical Flow", periodType="month",
     **{"from": "2025-01-01", "to": "2025-02-01", "countryKey": "DE", "adjacentSystemsKey": "DISTRIBUTION", "limit": 10})
show("points table", "Points", limit=3)
show("tariff/other", "operatorpointdirections", limit=3)
sys.exit(0)
