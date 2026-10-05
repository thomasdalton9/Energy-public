"""
Audit part 2 (West Europe gas borders): (a) own-side vs other-side per ordered border WITHOUT the master's fill rule, 2022-2025;
(b) ambiguous interconnection rows (one point/operator/direction listed with different far-side countries/types, where the master's
adjacency() keeps only the last); (c) flows of West-country operators that classify() drops (unclassified or same-country far side),
summed per point for 2023 and 2025 in TWh. Writes discovery_archive/europe/results/west_*.csv. Prints only.
"""
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "europe"))
sys.path.insert(0, ROOT)
import ENTSOG_GAS_FLOWS_DAILY as G  # noqa: E402

W = {"UK", "IE", "FR", "ES", "PT", "BE", "NL", "DE", "LU", "CH", "IT", "AT", "DK", "NO"}
OUT = os.path.join(HERE, "results")
os.makedirs(OUT, exist_ok=True)

ics = G.get_json("interconnections", {"limit": -1}).get("interconnections", [])
adj = G.adjacency()
label = {i["pointKey"]: i.get("pointLabel") for i in ics}
# (b) ambiguous rows
amb = {}
for i in ics:
    pk = i["pointKey"]
    if i.get("toOperatorKey"):
        amb.setdefault((pk, i["toOperatorKey"], "entry"), set()).add((i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey")))
    if i.get("fromOperatorKey"):
        amb.setdefault((pk, i["fromOperatorKey"], "exit"), set()).add((i.get("toInfrastructureTypeLabel"), i.get("toCountryKey")))
ambr = [(k, label.get(k[0]), sorted(map(str, v)), adj.get(k)) for k, v in amb.items() if len(v) > 1 and k[1][:2] in W]
print("AMBIGUOUS rows (point, label, alternatives, kept):", len(ambr))
for r in ambr:
    print(" ", r[0][1], r[0][2], r[1], r[2], "KEPT", r[3])

# (c) dropped flows, monthly windows
acc = {}
rows = []
for year in (2023, 2025):
    for m in range(1, 13):
        d0 = pd.Timestamp(year, m, 1)
        d1 = (d0 + pd.offsets.MonthEnd(0))
        data = G.get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0.date().isoformat(),
                                              "to": d1.date().isoformat(), "limit": -1}).get("operationalData", [])
        for r in data:
            v = r.get("value")
            op = r.get("operatorKey") or ""
            if v in (None, "") or op[:2] not in W:
                continue
            unit = r.get("unit", "kWh/d")
            f = {"kWh/d": 1e-6, "MWh/d": 1e-3, "GWh/d": 1.0}.get(unit)
            if f is None:
                continue
            g = float(v) * f
            c = G.classify(r, adj)
            a = adj.get((r["pointKey"], op, r["directionKey"]))
            if c is None:
                k = (year, op[:2], op, r["pointKey"], label.get(r["pointKey"]), r["directionKey"], str(a))
                acc[k] = acc.get(k, 0.0) + g / 1000
            elif c[1] in ("imports", "exports"):
                rows.append((pd.Timestamp(r["periodFrom"][:10]), c[0], r["directionKey"], c[2], r["pointKey"], op, g))
        print("done", year, m, flush=True)
dr = pd.DataFrame([dict(year=k[0], cc=k[1], operator=k[2], pointKey=k[3], point=k[4], dir=k[5], adj=k[6], TWh=round(v, 2)) for k, v in acc.items()])
dr.to_csv(os.path.join(OUT, "west_dropped_flows.csv"), index=False)
big = dr[(dr.TWh > 2) & ~dr.adj.str.contains("Distribution|Final Consumers|Storage|LNG|Production")].sort_values(["cc", "year", "TWh"], ascending=[True, True, False])
print("DROPPED (>2 TWh, not DSO/consumer/storage/LNG/production):")
for _, r in big.iterrows():
    print(f"  {r.year} {r.cc} {r.operator} {r.point} {r.dir} adj={r.adj} {r.TWh}")

# (a) own sides without fill
by_pt = {}
for day, c, d, oc, pk, op, g in rows:
    by_pt.setdefault((day.year, c, d, oc, pk), []).append(g)
side = {}
for (y, c, d, oc, pk), vals in by_pt.items():
    k = (y, c, d, oc, pk in G.VIP_KEYS)
    side[k] = side.get(k, 0.0) + G._dedupe_operators(vals)
own = {}
for (y, c, d, oc, vip), v in side.items():
    own[(y, c, d, oc)] = max(own.get((y, c, d, oc), 0.0), v) / 1000
recs = {}
for (y, c, d, oc), v in own.items():
    a, b = (c, oc) if d == "exit" else (oc, c)
    recs.setdefault((y, a, b), {})[d] = round(v, 1)
pd.DataFrame([dict(year=y, border=f"{a}>{b}", sender_exit=v.get("exit"), receiver_entry=v.get("entry")) for (y, a, b), v in sorted(recs.items())]).to_csv(
    os.path.join(OUT, "west_own_sides.csv"), index=False)
print("OWN SIDES 2025 (exit/entry, None = no row):")
for (y, a, b), v in sorted(recs.items()):
    if y == 2025 and max(v.get("exit", 0), v.get("entry", 0)) > 0.5:
        print(f"  {a}>{b} {v.get('exit')} / {v.get('entry')}")
