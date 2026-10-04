"""
Probe: one month of ENTSOG Physical Flow for chosen country pairs, point by point (reporting country, direction, point, operator, VIP flag, TWh),
with the old rule (own side summed) and the current border_flows rule side by side. Used to see why a country's balance moved when the border
rule changed (Bulgaria, Greece, Germany, Netherlands). Prints only.
Usage: PAIRS=BG-GR,BG-TR python3 ENTSOG_RULE_PROBE.py [from] [to]
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
import pandas as pd  # noqa: E402
import ENTSOG_GAS_FLOWS_DAILY as G  # noqa: E402

d0, d1 = (date.fromisoformat(a) for a in (sys.argv[1:3] if len(sys.argv) > 2 else ("2024-07-01", "2024-07-31")))
pairs = {frozenset(p.split("-")) for p in os.environ.get("PAIRS", "BG-GR,BG-TR,BG-RO,BG-RS,BG-MK").split(",")}
adj = G.adjacency()
data = G.get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0.isoformat(), "to": d1.isoformat(), "limit": -1}).get("operationalData", [])
rows, tab = [], {}
for r in data:
    v = r.get("value")
    if v in (None, ""):
        continue
    c = G.classify(r, adj)
    if c is None or c[1] not in ("imports", "exports"):
        continue
    cc, cat, oc = c
    if frozenset((cc, oc)) not in pairs:
        continue
    g = float(v) / 1e6
    d = "entry" if cat == "imports" else "exit"
    day = pd.Timestamp(r["periodFrom"][:10])
    rows.append((day, cc, d, oc, r["pointKey"], r.get("operatorKey"), g))
    k = (cc, d, oc, r.get("pointLabel"), r["pointKey"], r.get("operatorKey"), r["pointKey"] in G.VIP_KEYS)
    tab[k] = tab.get(k, 0.0) + g / 1000
print(f"{d0} -> {d1}: TWh per point/operator")
for k, v in sorted(tab.items()):
    print(" | ".join(str(x) for x in k), f"{v:.2f}")
fl = G.border_flows(rows)
tot = {}
for (day, a, b), (ex, im, mx) in fl.items():
    t = tot.setdefault((a, b), [0, 0, 0])
    t[0] += ex / 1000; t[1] += im / 1000; t[2] += mx / 1000
print("border: exports credited | imports credited | max  (TWh)")
for k, t in sorted(tot.items()):
    print(f"{k[0]}>{k[1]}", *[f"{x:.2f}" for x in t])
