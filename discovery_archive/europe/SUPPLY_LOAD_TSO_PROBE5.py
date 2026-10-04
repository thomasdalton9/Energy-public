"""Probe 5: Belgium monthly identity in Elia's own open data: generation (ods201) + net physical import (ods026) - total load (ods001), 2023-2025. Prints only."""
import signal
import sys
import urllib.parse

import requests

signal.signal(signal.SIGALRM, lambda *_: (_ for _ in ()).throw(TimeoutError("hard timeout")))
B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets"
H = {"User-Agent": "Mozilla/5.0"}


def q(ds, sel, where, group=None):
    p = {"select": sel, "where": where, "limit": 50}
    if group:
        p["group_by"] = group
    signal.alarm(60)
    try:
        r = requests.get(f"{B}/{ds}/records?" + urllib.parse.urlencode(p), headers=H, timeout=(10, 40))
        signal.alarm(0)
        return r.json()["results"] if r.ok else []
    except BaseException as e:  # noqa: BLE001
        signal.alarm(0)
        print("ERR", ds, where[:40], str(e)[:80], flush=True)
        return []


print("month | load | gen_total | pumped_gen | storage | net_import(-sum flows) | residual(gen+NI-load) | TWh")
for y in (2023, 2024, 2025):
    for m in range(1, 13):
        d0, d1 = f"{y}-{m:02d}-01", (f"{y}-{m + 1:02d}-01" if m < 12 else f"{y + 1}-01-01")
        w = f"datetime>=date'{d0}' and datetime<date'{d1}'"
        load = (q("ods001", "sum(totalload) as s", w) or [{}])[0].get("s") or 0
        gens = {r["fueltypeentsoe"]: r["s"] for r in q("ods201", "fueltypeentsoe, sum(generatedpower) as s", w, "fueltypeentsoe")}
        flows = {r["controlarea"]: r["s"] for r in q("ods026", "controlarea, sum(physicalflowatborder) as s", w, "controlarea")}
        gt = sum(gens.values()) / 4e6
        ni = -sum(v or 0 for v in flows.values()) / 4e6
        print(f"{y}-{m:02d} | {load / 4e6:6.3f} | {gt:6.3f} | {gens.get('Hydro Pumped Storage', 0) / 4e6:5.3f} | {gens.get('Energy Storage', 0) / 4e6:5.3f} | {ni:6.3f} | {gt + ni - load / 4e6:6.3f}", flush=True)
sys.exit(0)
