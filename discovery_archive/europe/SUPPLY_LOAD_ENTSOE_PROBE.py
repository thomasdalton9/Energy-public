"""
Probe: ENTSO-E physical (A11) vs finalised commercial schedules (A09) net imports for Serbia, Slovenia, Lithuania, Belgium, Spain,
TWh per year 2023-2025, per border. Prints only. (Supply > load check: do the flows explain it?)
"""
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE.replace("discovery_archive/europe", "europe"))
import entsoe_common as C  # noqa: E402
import ENTSOE_FLOWS_DAILY as F  # noqa: E402

TARGET = {"Serbia", "Slovenia", "Lithuania", "Belgium", "Spain"}


def annual(doc, a, b, y, eic):
    tot, days = 0.0, 0
    d1 = min(date(y + 1, 1, 1), C.today_utc().date())
    for w0, w1 in C.windows(C.utc_midnight(date(y, 1, 1)), C.utc_midnight(d1), 60):
        try:
            r = C.fetch_split({"documentType": doc, "out_Domain": eic[a], "in_Domain": eic[b]}, w0, w1, C.parse_energy, C.merge_energy)
        except Exception as e:  # noqa: BLE001
            return None
        if r:
            tot += sum(r["e"].values()) / 1e6
            days += len(r["e"])
    return tot, days


pairs, eic, country = F.borders()
net = {}
for a, b in pairs:
    ca, cb = country[a], country[b]
    if not ({ca, cb} & TARGET):
        continue
    for y in (2023, 2024, 2025):
        row = []
        for doc in ("A11", "A09"):
            fw = annual(doc, a, b, y, eic)
            bw = annual(doc, b, a, y, eic)
            row.append((fw, bw))
            for c, sgn in ((cb, 1), (ca, -1)):
                if c in TARGET:
                    net.setdefault((c, y, doc), 0.0)
                    net[(c, y, doc)] += sgn * ((fw[0] if fw else 0) - (bw[0] if bw else 0))
        f = lambda x: f"{x[0]:.2f}({x[1]}d)" if x else "n/a"
        print(f"{a}>{b} {y}: phys {f(row[0][0])}/{f(row[0][1])}  sched {f(row[1][0])}/{f(row[1][1])}", flush=True)
print("\nNET IMPORTS TWh (phys A11 | sched A09)")
for c in sorted(TARGET):
    for y in (2023, 2024, 2025):
        print(c, y, f"{net.get((c, y, 'A11'), 0):.2f} | {net.get((c, y, 'A09'), 0):.2f}")
