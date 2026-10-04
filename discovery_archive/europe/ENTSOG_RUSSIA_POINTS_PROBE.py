"""
Probe: Russian-border / Russian-route points in ENTSOG. Lists interconnections adjacent to RU/BY/UA/MD/TR or whose labels name a known Russian-route
point (Greifswald, Lubmin, Mallnow, Kondratki, Sudzha, Velke Kapusany, Strandzha, Isaccea, Imatra, Narva ...), pulls Physical Flow 2021-2025 per
point direction and prints annual TWh, the adjacency as the daily pull's classify() sees it, and how it is booked (country, category, origin) or
dropped. Prints only.
"""
import re
import sys
import time
import collections

import requests

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
LAND = re.compile(r"greifswald|lubmin|mallnow|kondratki|wysokoje|vysokoe|tietierowka|teterovka|drozdowicze|sudzha|sudza|velke kapusany|kapusany|"
                  r"uzhgorod|uzhhorod|beregovo|berehovo|isaccea|negru|strandzha|malkoclar|kipoi|kulata|sidirokastro|mediesu|kobrin|imatra|narva|luhamaa|"
                  r"kotlovka|kalvarija|nord ?stream|yamal|jamal|turkstream|eugal|opal|waidhaus|baumgarten|lanzhot|brandov|olbernhau|ceska|"
                  r"orlen|ukraine|russia|belarus|murfatlar|orlovka|tekovo|kiskundorozsma|beregdaroc|mosonmagyarovar|velke|grenzach", re.I)
ADJ_COUNTRIES = {"RU", "BY", "UA", "MD", "TR"}


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
cands, adj = {}, {}
for i in ics:
    pk = i["pointKey"]
    if i.get("toOperatorKey"):
        adj[(pk, i["toOperatorKey"], "entry")] = (i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"))
    if i.get("fromOperatorKey"):
        adj[(pk, i["fromOperatorKey"], "exit")] = (i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"))
    labels = " ".join(str(i.get(k) or "") for k in ("pointLabel", "fromPointLabel", "toPointLabel", "fromOperatorLabel", "toOperatorLabel"))
    ra = i.get("fromCountryKey") in ADJ_COUNTRIES or i.get("toCountryKey") in ADJ_COUNTRIES
    if not (ra or LAND.search(labels)):
        continue
    for opk, d in ((i.get("toOperatorKey"), "entry"), (i.get("fromOperatorKey"), "exit")):
        if opk:
            cands[(opk, pk, d)] = i.get("pointLabel")
print("candidate point-directions:", len(cands))


def booked(opk, pk, d):
    a = adj.get((pk, opk, d))
    country = opk[:2]
    if a is None:
        return "DROPPED (no adjacency)"
    typ, ac = a
    if typ == "Transmission":
        if ac and ac != country:
            return f"{country} {'imports' if d == 'entry' else 'exports'} origin/dest {ac}"
        return f"DROPPED (same country {ac})"
    return f"adj {typ}/{ac}: " + ("booked non-border" if typ in ("Production", "LNG Terminals", "Storage", "Distribution", "Final Consumers") else "DROPPED")


rows = []
for (opk, pk, d), lab in sorted(cands.items()):
    data = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": "2021-01-01", "to": "2025-12-31",
                                   "pointDirection": f"{opk}{pk}{d}", "limit": -1}).get("operationalData", [])
    yr = collections.defaultdict(float)
    for r in data:
        try:
            v = float(r["value"])
        except (TypeError, ValueError, KeyError):
            continue
        yr[r["periodFrom"][:4]] += v / 1e9
    if any(yr.values()):
        rows.append((opk, pk, lab, d, adj.get((pk, opk, d)), booked(opk, pk, d), [round(yr.get(str(y), 0), 1) for y in range(2021, 2026)]))
print("operator | pointKey | label | dir | adjacency | booking | TWh 2021..2025")
for r in sorted(rows, key=lambda x: -sum(x[6])):
    print(" | ".join(str(x) for x in r))
sys.exit(0)
