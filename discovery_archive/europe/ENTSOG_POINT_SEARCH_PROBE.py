"""
Probe: per-point ENTSOG Physical Flow totals (TWh per year) for every (operator, point, direction) whose operator key starts with one of the
OPS prefixes (comma list, e.g. "AT-,UK-") or whose point label / adjacent operator matches the LABELS regex (e.g. "Haidach|Baumgarten"),
with the adjacent system. Finds points the country-balance classification drops (storage, production, terminals) and missing ends.
Prints only. Env: OPS, LABELS, YEARS (default 2022,2023,2024,2025).
"""
import os
import re
import sys
import time

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
OPS = tuple(o for o in os.environ.get("OPS", "").split(",") if o)
LABELS = re.compile(os.environ.get("LABELS", "$^"), re.I)
YEARS = [int(y) for y in os.environ.get("YEARS", "2022,2023,2024,2025").split(",")]


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
cand = {}
for i in ics:
    lab = str(i.get("pointLabel") or "")
    for opk, d, adjt, adjc, adjo in ((i.get("toOperatorKey"), "entry", i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"), i.get("fromOperatorKey")),
                                    (i.get("fromOperatorKey"), "exit", i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"), i.get("toOperatorKey"))):
        if not opk:
            continue
        if (OPS and opk.startswith(OPS)) or LABELS.search(lab):
            cand[(opk, i["pointKey"], d)] = (lab, adjt, adjc, adjo)
print("candidates:", len(cand), flush=True)
print("operator | dir | point | key | adj type | adj country | adj operator | " + " | ".join(map(str, YEARS)) + " (TWh)", flush=True)
for (opk, pk, d), (lab, adjt, adjc, adjo) in sorted(cand.items()):
    tot = []
    for y in YEARS:
        data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                       "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
        tot.append(sum(float(r["value"]) for r in data if r.get("value") not in (None, "")) / 1e9)
    if any(abs(t) > 0.05 for t in tot):
        print(" | ".join([opk, d, lab[:45], pk, str(adjt), str(adjc), str(adjo)] + [f"{t:.1f}" for t in tot]), flush=True)
sys.exit(0)
