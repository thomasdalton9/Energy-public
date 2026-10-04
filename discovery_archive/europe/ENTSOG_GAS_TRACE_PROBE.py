"""
Probe: every ENTSOG Physical Flow point touching GB/UK, FR, HU or GR (as operator country or adjacent country), with 2023 and 2025 TWh,
how ENTSOG_GAS_FLOWS_DAILY.classify() treats it (category / origin or UNCLASSIFIED), to find missing or double-counted points
behind the gas balance errors (UK +6%, FR +7%, HU +14%, GR -15%). Prints only.
"""
import sys
import time
from collections import defaultdict

import requests

sys.path.insert(0, "europe")
API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
TARGET = {"UK", "GB", "FR", "HU", "GR"}


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
adj = {}
for i in ics:
    pk = i["pointKey"]
    if i.get("toOperatorKey"):
        adj[(pk, i["toOperatorKey"], "entry")] = (i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"), i.get("fromOperatorKey"))
    if i.get("fromOperatorKey"):
        adj[(pk, i["fromOperatorKey"], "exit")] = (i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"), i.get("toOperatorKey"))
tot = defaultdict(lambda: defaultdict(float))
meta = {}
for y in (2023, 2025):
    for m in range(1, 13):
        d0 = f"{y}-{m:02d}-01"
        d1 = f"{y + (m == 12)}-{(m % 12) + 1:02d}-01"
        import datetime as dt
        d1 = (dt.date.fromisoformat(d1) - dt.timedelta(days=1)).isoformat()
        rows = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0, "to": d1, "limit": -1}).get("operationalData", [])
        print(y, m, len(rows), flush=True)
        for r in rows:
            v = r.get("value")
            if v in (None, ""):
                continue
            op, pk, d = r.get("operatorKey") or "", r["pointKey"], r["directionKey"]
            a = adj.get((pk, op, d), (None, None, None))
            if not ({op[:2], a[1]} & TARGET):
                continue
            k = (op, pk, d)
            tot[k][y] += float(v) / 1e9
            meta[k] = (r.get("pointLabel"), a)
print("operator | point | dir | label | adjacent type/country/operator | TWh 2023 | TWh 2025")
for k in sorted(tot, key=lambda k: (k[0][:2], k[2], -tot[k][2025])):
    if max(tot[k][2023], tot[k][2025]) < 0.3:
        continue
    lab, a = meta[k]
    print(f"{k[0]} | {k[1]} | {k[2]} | {lab} | {a} | {tot[k][2023]:.1f} | {tot[k][2025]:.1f}")
