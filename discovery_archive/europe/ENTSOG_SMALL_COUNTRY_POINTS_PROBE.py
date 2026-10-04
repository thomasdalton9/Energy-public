"""
Probe (small-country gas balances; countries via argv): per-point ENTSOG Physical Flow (TWh) for 2023, 2024, 2025 at every interconnection point of CZ and AT operators, entry and exit,
with the adjacent system, to find why Czech imports double in 2025 (ENTSOG 185 TWh vs NET4GAS border entry 93) and why Austria's
balance is far off in 2023-24. Prints only.
"""
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
COUNTRIES = tuple((sys.argv[1] if len(sys.argv) > 1 else "LU,EE,LV,BG,BE").split(","))


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
cands = {}
for i in ics:
    for opk, d, adjt, adjc, adjop in ((i.get("toOperatorKey"), "entry", i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"), i.get("fromOperatorKey")),
                                      (i.get("fromOperatorKey"), "exit", i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"), i.get("toOperatorKey"))):
        if opk and opk[:2] in COUNTRIES:
            cands[(opk, i["pointKey"], d)] = (i.get("pointLabel"), adjt, adjc, adjop)
print("candidates:", len(cands), flush=True)
rows = []
for (opk, pk, d), (lab, adjt, adjc, adjop) in sorted(cands.items()):
    out = []
    for y in (2022, 2023, 2024, 2025):
        data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                       "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
        out.append(sum(float(r["value"]) for r in data if r.get("value") not in (None, "")) / 1e9)
    if any(abs(x) > 0.05 for x in out):
        rows.append((opk[:2], d, lab, opk, pk, adjt, adjc, adjop, *[round(x, 1) for x in out]))
print("cc | dir | point | operator | key | adjacent type | adj country | adj operator | 2022 | 2023 | 2024 | 2025 (TWh)")
for r in sorted(rows, key=lambda x: (x[0], x[1], -x[11])):
    print(" | ".join(str(x) for x in r))
sys.exit(0)
