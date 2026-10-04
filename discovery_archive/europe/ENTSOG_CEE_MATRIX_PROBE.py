"""
Probe: border matrix for central/eastern/south-east Europe, Nordics and Baltics. For each country pair (env PAIRS 'PL-DE,PL-CZ', unordered) lists every
ENTSOG interconnection point in both directions (sender exit, receiver entry, every reporting operator) with the Physical Flow in TWh for each year,
then a summary per direction: sender-side sum, receiver-side sum (plain sums of all rows, so VIP + physical points double count - read the point rows).
Prints only.
"""
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
PAIRS = [tuple(x.split('-')) for x in os.environ['PAIRS'].split(',')]
YEARS = [int(y) for y in os.environ.get('YEARS', '2022,2023,2024,2025').split(',')]


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


def year_tot(pd_key, y):
    data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                   "pointDirection": pd_key, "limit": -1}).get("operationalData", [])
    return sum(float(r["value"]) for r in data if r.get("value") not in (None, "")) / 1e9


ics = get("interconnections", {"limit": -1}).get("interconnections", [])
print("interconnections:", len(ics), flush=True)
for a, b in PAIRS:
    print("=" * 100, f"\n{a} <-> {b}", flush=True)
    jobs, seen = [], set()
    for i in ics:
        fc, tc = i.get("fromCountryKey"), i.get("toCountryKey")
        if {fc, tc} != {a, b}:
            continue
        for opk, d in ((i.get("toOperatorKey"), "entry"), (i.get("fromOperatorKey"), "exit")):
            if not opk:
                continue
            key = (opk, i["pointKey"], d)
            if key in seen:
                continue
            seen.add(key)
            jobs.append((opk, d, i.get("pointLabel"), i["pointKey"], fc, tc))
    def run(j):
        return [year_tot(f"{j[0]}{j[3]}{j[1]}", y) for y in YEARS]
    with ThreadPoolExecutor(6) as ex:
        res = list(ex.map(run, jobs))
    print("rows: op | dir | point | from>to | " + " ".join(map(str, YEARS)))
    flows = {}
    for j, t in sorted(zip(jobs, res), key=lambda x: (x[0][4], x[0][1], -x[1][-1])):
        print(f"{j[0]} | {j[1]} | {j[2]} | {j[4]}>{j[5]} | " + " ".join(f"{x:.1f}" for x in t), flush=True)
        # physical direction: exit of op country means gas leaves op country towards the other
        opc = j[0][:2]
        frm, to = (opc, b if opc == a else a) if j[1] == "exit" else ((b if opc == a else a), opc)
        for k, x in enumerate(t):
            flows.setdefault((frm, to, 'exit' if j[1] == 'exit' else 'entry'), [0] * len(YEARS))[k] += x
    print("summary (plain sums, may double count VIPs):")
    for k, v in sorted(flows.items()):
        print(f"  {k[0]}>{k[1]} {'sender exit' if k[2]=='exit' else 'receiver entry'}: " + " ".join(f"{x:.1f}" for x in v))
sys.exit(0)
