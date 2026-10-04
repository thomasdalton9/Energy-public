"""
Probe: both ends of the UK-continent interconnectors and the German borders. For each pair of countries lists every interconnection point,
the reporting operator on each side and direction, and the Physical Flow in TWh for 2023, 2024, 2025, to find missing ends (e.g. the Dutch
end of BBL), duplicated points (a physical point and a virtual one) and the jump in German exports in 2025. Prints only.
"""
import os
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
PAIRS = [("UK", "BE"), ("UK", "NL"), ("UK", "FR"), ("UK", "IE"), ("UK", "DE"), ("DE", "CZ"), ("DE", "AT"), ("DE", "PL"), ("DE", "NL"),
         ("DE", "BE"), ("DE", "FR"), ("DE", "CH"), ("DE", "DK"), ("DE", "LU")]

if os.environ.get('PAIRS'):   # e.g. 'AT-SK,AT-HU'
    PAIRS = [tuple(x.split('-')) for x in os.environ['PAIRS'].split(',')]
YEARS = [int(y) for y in os.environ.get('YEARS', '2023,2024,2025').split(',')]


def get(path, params, tries=4):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(8 * (i + 1))
        except requests.RequestException:
            time.sleep(8 * (i + 1))
    return {}


ics = get("interconnections", {"limit": -1}).get("interconnections", [])
seen = set()
for a, b in PAIRS:
    print("=" * 100, f"\n{a} <-> {b}", flush=True)
    rows = []
    for i in ics:
        fc, tc = i.get("fromCountryKey"), i.get("toCountryKey")
        if {fc, tc} != {a, b}:
            continue
        for opk, d in ((i.get("toOperatorKey"), "entry"), (i.get("fromOperatorKey"), "exit")):
            if not opk or opk[:2] not in (a, b):
                continue
            key = (opk, i["pointKey"], d)
            if key in seen:
                continue
            seen.add(key)
            tot = []
            for y in YEARS:
                data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                               "pointDirection": f"{opk}{i['pointKey']}{d}", "limit": -1}).get("operationalData", [])
                tot.append(sum(float(r["value"]) for r in data if r.get("value") not in (None, "")) / 1e9)
            rows.append((opk[:2], d, i.get("pointLabel"), opk, i["pointKey"], *[round(x, 1) for x in tot]))
    for r in sorted(rows, key=lambda x: (x[0], x[1], -x[5])):
        print(" | ".join(str(x) for x in r))
sys.exit(0)
