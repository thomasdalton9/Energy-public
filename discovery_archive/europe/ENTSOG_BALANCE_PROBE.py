"""
Probe for a country-level gas balance from ENTSOG physical flows: does one unfiltered request return every point for a
day or a month, what does a row look like, how big/slow is it, and how do rows join to the interconnections list
(from/to country, infrastructure type) so flows can be classed as pipeline import / export, LNG, production, storage,
distribution / final-consumer demand?

Usage: python3 ENTSOG_BALANCE_PROBE.py
"""
import json
import time
from collections import Counter

import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
API = "https://transparency.entsog.eu/api/v1"


def get(url, **params):
    t0 = time.time()
    r = requests.get(url, params=params, headers=H, timeout=(15, 300))
    return r, time.time() - t0


def main():
    r, dt = get(f"{API}/interconnections", limit=-1)
    ics = r.json()["interconnections"]
    print(f"interconnections: {r.status_code} {len(ics)} rows {dt:.1f}s")
    print("keys:", sorted(ics[0]))
    print("sample:", json.dumps(ics[5])[:900])
    for days, label in ((1, "1 day"), (7, "7 days"), (31, "31 days")):
        r, dt = get(f"{API}/operationalData", indicator="Physical Flow", periodType="day",
                    **{"from": "2026-08-01", "to": f"2026-08-{days:02d}" if days < 31 else "2026-08-31"}, limit=-1)
        print(f"\noperationalData unfiltered {label}: HTTP {r.status_code} {len(r.content) / 1e6:.1f} MB {dt:.1f}s")
        if r.status_code != 200:
            print(r.text[:300])
            continue
        rows = r.json().get("operationalData", [])
        print("rows:", len(rows), "| distinct points:", len({(x["pointKey"], x["directionKey"]) for x in rows}),
              "| distinct days:", len({x["periodFrom"][:10] for x in rows}))
        if label == "1 day":
            print("row keys:", sorted(rows[0]))
            for x in rows[:3]:
                print({k: x[k] for k in ("pointKey", "pointLabel", "operatorKey", "directionKey", "value", "unit", "periodFrom",
                                         "adjacentSystemsLabel", "adjacentSystemsKey") if k in x})
            print("directionKey:", Counter(x["directionKey"] for x in rows))
            print("units:", Counter(x["unit"] for x in rows))
            print("flow value empty:", sum(1 for x in rows if x["value"] in (None, "")))
            # join to interconnections
            idx = {}
            for i in ics:
                idx.setdefault(i["pointKey"], []).append(i)
            miss = [x["pointKey"] for x in rows if x["pointKey"] not in idx]
            print("rows whose point is not in interconnections:", len(miss), sorted(set(miss))[:8])
            # classify
            by = Counter()
            for x in rows:
                for i in idx.get(x["pointKey"], [])[:1]:
                    by[(x["directionKey"], i.get("fromInfrastructureTypeLabel"), i.get("toInfrastructureTypeLabel"))] += 1
            for k, v in by.most_common(14):
                print("  ", k, v)
    # one well-known point both ways, to confirm entry/exit semantics: Moffat (IE) entry and a DE-DK border
    for pk in ("ITP-00495",):
        r, dt = get(f"{API}/operationalData", indicator="Physical Flow", periodType="day", pointKey=pk,
                    **{"from": "2026-09-01", "to": "2026-09-03"}, limit=-1)
        print("\npoint", pk, r.status_code, [(x["directionKey"], x["value"], x["unit"]) for x in r.json().get("operationalData", [])][:6])
    # 5-year depth check
    r, dt = get(f"{API}/operationalData", indicator="Physical Flow", periodType="day",
                **{"from": "2021-10-04", "to": "2021-10-04"}, limit=-1)
    print("\noldest day check 2021-10-04:", r.status_code, len(r.json().get("operationalData", [])) if r.status_code == 200 else r.text[:200])


if __name__ == "__main__":
    main()
