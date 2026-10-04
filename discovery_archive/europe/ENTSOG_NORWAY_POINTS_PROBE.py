"""
Probe: which ENTSOG points carry Norwegian gas, how are they classified, and what do they sum to (2025 TWh) per receiving country.
Lists interconnections whose adjacent system is in Norway (or whose label names a known Norwegian landfall), then pulls the 2025 Physical
Flow for each and prints annual TWh by (reporting country, point, direction, adjacent type/country). Compared with Gassco's 2025 flows
(DE 649, FR 167, BE 171, GB 243, other 45 TWh from norway_gassco_gas_flows_daily.xlsx). Prints only.
"""
import re
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
LAND = re.compile(r"emden|dornum|dunkerque|dunkirk|zeebrugge|easington|st\.? ?fergus|nybro|den helder|bacton|balgzand|"
                  r"norpipe|europipe|franpipe|zeepipe|vesterled|langeled|statpipe|baltic ?pipe|kollsnes|karst|kaarst|"
                  r"draupner|ekofisk|tyra|nyhamna|aukra|heimdal|sleipner|troll|oseberg|nybro|norway|norwegian|gassco", re.I)


def get(path, params, tries=4):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(10 * (i + 1))
        except requests.RequestException:
            time.sleep(10 * (i + 1))
    return {}


ics = get("interconnections", {"limit": -1}).get("interconnections", [])
print("interconnections:", len(ics))
cands = {}
for i in ics:
    labels = " ".join(str(i.get(k) or "") for k in ("pointLabel", "fromPointLabel", "toPointLabel", "fromOperatorLabel", "toOperatorLabel"))
    norway_adj = i.get("fromCountryKey") == "NO" or i.get("toCountryKey") == "NO"
    if not (norway_adj or LAND.search(labels)):
        continue
    for opk, d, adjt, adjc in ((i.get("toOperatorKey"), "entry", i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey")),
                               (i.get("fromOperatorKey"), "exit", i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"))):
        if opk:
            cands[(opk, i["pointKey"], d)] = (i.get("pointLabel"), adjt, adjc, bool(norway_adj))
print("candidate point-directions:", len(cands))
rows = []
for (opk, pk, d), (lab, adjt, adjc, nadj) in sorted(cands.items()):
    data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": "2025-01-01", "to": "2025-12-31",
                                   "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
    tot, n = 0.0, 0
    for r in data:
        try:
            tot += float(r["value"]) / 1e9      # kWh/d -> TWh
            n += 1
        except (TypeError, ValueError, KeyError):
            pass
    rows.append((opk[:2], lab, opk, pk, d, adjt, adjc, nadj, round(tot, 1), n))
print("country | point | operator | key | dir | adjacent type | adjacent country | norway-adjacent | 2025 TWh | days")
for r in sorted(rows, key=lambda x: (-x[8])):
    if r[8] or r[7]:
        print(" | ".join(str(x) for x in r))
agg = {}
for r in rows:
    if r[5] == "Transmission" and r[7]:
        agg[(r[0], r[4], r[6])] = agg.get((r[0], r[4], r[6]), 0) + r[8]
print("Norway-adjacent transmission sums by (reporting country, direction, adjacent):", agg)
sys.exit(0)
