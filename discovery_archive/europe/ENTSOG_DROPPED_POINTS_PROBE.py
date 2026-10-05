"""
Probe: ENTSOG points the country classification (ENTSOG_GAS_FLOWS_DAILY.classify) drops for the CEE / Nordic / Baltic operators: rows whose adjacent
system is not a known type or is a Transmission system of the same country (or has no country). Prints the physical flow in TWh for 2022-2025 of
each such operator/point/direction so large dropped flows (Kondratki, Emden-style cases) show up. Also prints the adjacent-type labels seen.
"""
import os
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
CC = os.environ.get('CC', 'PL,CZ,SK,HU,RO,BG,GR,HR,SI,LT,LV,EE,FI,SE,DK,AT,RS,MK,BA,AL,MD,UA,TR').split(',')
YEARS = [2022, 2023, 2024, 2025]


def get(path, params, tries=4):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(6 * (i + 1))
        except requests.RequestException:
            time.sleep(6 * (i + 1))
    return {}


ics = get("interconnections", {"limit": -1}).get("interconnections", [])
types = Counter()
jobs, seen = [], set()
for i in ics:
    for opk, d, typ, acc in ((i.get("toOperatorKey"), "entry", i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey")),
                             (i.get("fromOperatorKey"), "exit", i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"))):
        if not opk or opk[:2] not in CC:
            continue
        types[typ] += 1
        ok = typ in ("Production", "LNG Terminals", "Storage", "Distribution", "Final Consumers") or (typ == "Transmission" and acc and acc != opk[:2])
        if typ == "Transmission" and acc and acc != opk[:2]:
            ok = True
        key = (opk, i["pointKey"], d)
        if ok or key in seen:
            continue
        seen.add(key)
        jobs.append((opk, i["pointKey"], d, i.get("pointLabel"), typ, acc, i.get("fromCountryKey"), i.get("toCountryKey")))
print("adjacent type labels:", dict(types), "candidates:", len(jobs), flush=True)


def run(j):
    out = []
    for y in YEARS:
        data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                       "pointDirection": f"{j[0]}{j[1]}{j[2]}", "limit": -1}).get("operationalData", [])
        out.append(sum(float(r["value"]) for r in data if r.get("value") not in (None, "")) / 1e9)
    return out


with ThreadPoolExecutor(8) as ex:
    res = list(ex.map(run, jobs))
for j, t in sorted(zip(jobs, res), key=lambda x: -max(x[1])):
    if max(t) >= 0.3:
        print(f"{j[0]} | {j[2]} | {j[3]} | adj={j[4]}/{j[5]} from={j[6]} to={j[7]} | " + " ".join(f"{x:.1f}" for x in t))
sys.exit(0)
