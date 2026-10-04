"""
Probe 2: Emden (EPT1, EPT2, NPT) and Dornum have no 'Physical Flow' rows in ENTSOG. Check the other indicators (Allocation, Nomination,
Renomination, GCV availability) for every Norway-adjacent point in DE/NL and print 2025 TWh per operator/direction/indicator, plus the list of
indicators ENTSOG offers for a sample point. Prints only.
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


POINTS = [("DE-TSO-0002", "ITP-00209"), ("DE-TSO-0002", "ITP-00210"), ("DE-TSO-0005", "ITP-00208"), ("DE-TSO-0005", "ITP-00209"),
          ("DE-TSO-0005", "ITP-00210"), ("DE-TSO-0009", "ITP-00208"), ("DE-TSO-0009", "ITP-00209"), ("DE-TSO-0009", "ITP-00525"),
          ("NL-TSO-0001", "ITP-00209"), ("NL-TSO-0001", "ITP-00210"), ("DE-TSO-0009", "ITP-00210")]
INDICATORS = ["Physical Flow", "Allocation", "Nomination", "Renomination", "Firm Technical", "Interruptible Total"]
print("indicators offered:", [i.get("indicator") for i in get("indicators", {"limit": 50}).get("indicators", [])][:40])
for opk, pk in POINTS:
    for d in ("entry", "exit"):
        out = []
        for ind in INDICATORS[:4]:
            data = get("operationalData", {"indicator": ind, "periodType": "day", "from": "2025-01-01", "to": "2025-12-31",
                                           "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
            tot, n = 0.0, 0
            for r in data:
                try:
                    tot += float(r["value"]) / 1e9
                    n += 1
                except (TypeError, ValueError, KeyError):
                    pass
            out.append(f"{ind}: {tot:.1f} TWh/{n}d")
        print(opk, pk, d, " | ".join(out), flush=True)
# every point whose label contains Emden/Dornum/Norpipe/Europipe, any operator, in the interconnections list
ics = get("interconnections", {"limit": -1}).get("interconnections", [])
seen = set()
for i in ics:
    lab = str(i.get("pointLabel") or "")
    if any(k in lab.lower() for k in ("emden", "dornum", "norpipe", "europipe", "statpipe")):
        key = (lab, i.get("fromOperatorKey"), i.get("toOperatorKey"), i.get("fromCountryKey"), i.get("toCountryKey"))
        if key not in seen:
            seen.add(key)
            print("IC:", key)
sys.exit(0)
