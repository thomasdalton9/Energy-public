"""
Probe 2: catalogue search for Nord Stream / Yamal / Belarus-Poland points in ENTSOG (operatorpointdirections + interconnections), listing
every point whose label matches, with annual TWh 2021-2025 for Physical Flow, Allocation and Nomination, to see whether ENTSOG publishes
them at all. Prints only.
"""
import re
import sys
import time
import collections

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
LAB = re.compile(r"greifswald|lubmin|nord ?stream|\bnel\b|\bopal\b|wysokoje|vysokoe|kondratki|tietierowka|tieterowka|drozdowicze|brest|kobrin|"
                 r"mallnow|yamal|jamal|belarus|kaliningrad|emden|dornum|ostsee|baltic", re.I)


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


opd = get("operatorpointdirections", {"limit": -1}).get("operatorpointdirections", [])
print("operatorpointdirections:", len(opd), flush=True)
if opd:
    print("keys:", sorted(opd[0].keys()), flush=True)
seen = {}
for r in opd:
    lab = " ".join(str(r.get(k) or "") for k in ("pointLabel", "operatorLabel", "adjacentSystemsLabel", "tSOCountry", "pointTpMapping"))
    if LAB.search(lab):
        seen[(r["operatorKey"], r["pointKey"], r["directionKey"])] = (r.get("pointLabel"), r.get("tSOCountry"), r.get("adjacentSystemsLabel"), r.get("pointType"))
print("matching:", len(seen), flush=True)
for (opk, pk, d), meta in sorted(seen.items()):
    out = []
    for ind in ("Physical Flow", "Allocation", "Nomination"):
        data = get("operationalData", {"indicator": ind, "periodType": "day", "from": "2021-01-01", "to": "2025-12-31",
                                       "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
        yr = collections.defaultdict(float)
        for x in data:
            try:
                yr[x["periodFrom"][:4]] += float(x["value"]) / 1e9
            except (TypeError, ValueError, KeyError):
                pass
        out.append((ind[:4], [round(yr.get(str(y), 0), 1) for y in range(2021, 2026)]))
    if any(any(v) for _, v in out) or re.search(r"greifswald|lubmin|nord|kondratki|wysokoje|tietierowka|drozdowicze|kobrin", str(meta), re.I):
        print(opk, pk, d, meta, out, flush=True)
sys.exit(0)
