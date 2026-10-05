"""
Audit: every cross-border gas interconnection touching Western Europe (UK IE FR ES PT BE NL DE LU CH IT AT DK NO), both sides, 2022-2025.
Pulls ENTSOG Physical Flow (daily, per point/operator/direction), then applies the master's own border_flows() de-duplication
(operators once, VIP vs physical = larger-of) and writes
  discovery_archive/europe/results/west_borders_matrix.csv  (ordered border, year: sender exit TWh, receiver entry TWh, used)
  discovery_archive/europe/results/west_borders_points.csv  (point-level TWh per year, operator, VIP flag)
Prints a compact matrix with flags. Prints/writes only.
"""
import os
import sys
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "europe"))
sys.path.insert(0, ROOT)
import ENTSOG_GAS_FLOWS_DAILY as G  # noqa: E402

W = set(os.environ.get("COUNTRIES", "UK,IE,FR,ES,PT,BE,NL,DE,LU,CH,IT,AT,DK,NO").split(","))
YEARS = [2022, 2023, 2024, 2025]
OUT = os.path.join(HERE, "results")
os.makedirs(OUT, exist_ok=True)

ics = G.get_json("interconnections", {"limit": -1}).get("interconnections", [])
G.adjacency()   # fills VIP_KEYS
todo = {}       # (opKey, pointKey, dir) -> (reportingCountry, otherCountry, label, otherOperator)
for i in ics:
    fc, tc = i.get("fromCountryKey"), i.get("toCountryKey")
    if not fc or not tc or fc == tc or not ({fc, tc} & W):
        continue
    if i.get("toInfrastructureTypeLabel") not in ("Transmission", None) and i.get("fromInfrastructureTypeLabel") not in ("Transmission", None):
        pass
    if i.get("fromInfrastructureTypeLabel") != "Transmission" or i.get("toInfrastructureTypeLabel") != "Transmission":
        continue
    if i.get("fromOperatorKey"):
        todo[(i["fromOperatorKey"], i["pointKey"], "exit")] = (fc, tc, i.get("pointLabel"))
    if i.get("toOperatorKey"):
        todo[(i["toOperatorKey"], i["pointKey"], "entry")] = (tc, fc, i.get("pointLabel"))
print(f"{len(todo)} point-operator-directions on {len({(v[0], v[1]) for v in todo.values()})} country pairs", flush=True)


def pull(k):
    op, pk, d = k
    data = G.get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": "2022-01-01", "to": "2025-12-31",
                                          "pointDirection": f"{op}{pk}{d}", "limit": -1}).get("operationalData", [])
    out = []
    for r in data:
        v = r.get("value")
        if v in (None, ""):
            continue
        unit = r.get("unit", "kWh/d")
        f = {"kWh/d": 1e-6, "MWh/d": 1e-3, "GWh/d": 1.0}.get(unit)
        if f is None:
            continue
        out.append((pd.Timestamp(r["periodFrom"][:10]), float(v) * f))
    return k, out


with ThreadPoolExecutor(4) as ex:
    res = dict(ex.map(pull, todo))

rows, pts = [], []
for (op, pk, d), series in res.items():
    c, oc, label = todo[(op, pk, d)]
    if op[:2] != c:
        continue
    for day, g in series:
        rows.append((day, c, d, oc, pk, op, g))
    tot = {y: sum(g for day, g in series if day.year == y) / 1000 for y in YEARS}
    pts.append(dict(country=c, other=oc, dir=d, point=label, pointKey=pk, operator=op, vip=pk in G.VIP_KEYS, **{str(y): round(t, 2) for y, t in tot.items()}))
pd.DataFrame(pts).sort_values(["country", "other", "dir", "2025"], ascending=[True, True, True, False]).to_csv(os.path.join(OUT, "west_borders_points.csv"), index=False)

bf = G.border_flows(rows)
m = {}
for (day, a, b), (exv, imv, big) in bf.items():
    for key, v in (("sender_exit", exv), ("receiver_entry", imv), ("used", big)):
        m[(a, b, day.year, key)] = m.get((a, b, day.year, key), 0.0) + v / 1000
# raw (own side only, no fill) for transparency
raw = {}
for (day, c, d, oc, pk, op, g) in rows:
    pass
recs = []
for (a, b) in sorted({(a, b) for (a, b, y, k) in m}):
    rec = {"border": f"{a}>{b}"}
    for y in YEARS:
        for k in ("sender_exit", "receiver_entry", "used"):
            rec[f"{y}_{k}"] = round(m.get((a, b, y, k), 0.0), 1)
    recs.append(rec)
df = pd.DataFrame(recs)
df.to_csv(os.path.join(OUT, "west_borders_matrix.csv"), index=False)
print("border | 2022 ex/en | 2023 | 2024 | 2025 | FLAG")
for r in recs:
    flags = []
    for y in YEARS:
        e, i = r[f"{y}_sender_exit"], r[f"{y}_receiver_entry"]
        if max(e, i) < 0.5:
            continue
        if min(e, i) == 0 and max(e, i) > 0.5:
            flags.append(f"{y}:ONE-SIDE-ZERO")
        elif abs(e - i) > 2 and abs(e - i) / max(e, i) > 0.10:
            flags.append(f"{y}:DIFF")
    if max(r[f"{y}_used"] for y in YEARS) < 0.5:
        continue
    print(r["border"], "|", " | ".join(f"{r[f'{y}_sender_exit']}/{r[f'{y}_receiver_entry']}" for y in YEARS), "|", " ".join(flags))
# points with no data at all
print("NO DATA points:", sorted({(p['country'], p['other'], p['point'], p['dir']) for p in pts if all(p[str(y)] == 0 for y in YEARS)}))
