"""
Probe: ENTSOG storage points whose far side or operator sits in a different country than the storage's own country (candidates for storage booked
in one AGSI+ zone but connected to another country's grid). Prints interconnections touching a storage (infrastructure type contains 'Storage' or
point key starts UGS) with from/to country + operator, then 2025 physical-flow totals (TWh) for each of those pointDirections.
"""
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}


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
print("interconnections", len(ics), "keys", sorted(ics[0].keys()) if ics else None, flush=True)
cands = []
for i in ics:
    st = "storage" in (str(i.get("fromInfrastructureTypeLabel")) + str(i.get("toInfrastructureTypeLabel"))).lower() or str(i.get("pointKey", "")).startswith("UGS")
    if not st:
        continue
    fc, tc = i.get("fromCountryKey"), i.get("toCountryKey")
    print("IC", i.get("pointKey"), "|", i.get("pointLabel"), "| from", i.get("fromOperatorKey"), fc, i.get("fromInfrastructureTypeLabel"), "| to", i.get("toOperatorKey"), tc, i.get("toInfrastructureTypeLabel"), "| CROSS" if fc != tc else "", flush=True)
    if fc != tc:
        cands.append(i)
print("== cross-country storage links:", len(cands), flush=True)
for i in cands:
    for opk, d in ((i.get("toOperatorKey"), "entry"), (i.get("fromOperatorKey"), "exit")):
        tot = {}
        for yr in (2023, 2024, 2025):
            rows = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{yr}-01-01", "to": f"{yr}-12-31",
                                           "pointDirection": f"{opk}{i['pointKey']}{d}", "limit": -1}).get("operationalData", [])
            tot[yr] = sum(float(r["value"]) for r in rows if r.get("value") not in (None, "")) / 1e9
        if any(tot.values()):
            print("FLOW", i["pointKey"], i.get("pointLabel"), opk, d, {k: round(v, 1) for k, v in tot.items()}, "from", i.get("fromCountryKey"), "to", i.get("toCountryKey"), flush=True)
sys.exit(0)
