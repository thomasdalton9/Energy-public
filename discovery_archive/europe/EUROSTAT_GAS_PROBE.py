"""One-off probe: Eurostat monthly natural gas supply and consumption (nrg_cb_gasm) - dimension codes (balance items, units),
EU27 countries covered, latest month, and a sample (Germany, 2025) so we can use it as the second consumption source next to
ENTSOG. Prints only."""
import json
import sys

import requests

B = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm"
r = requests.get(B, params={"format": "JSON", "lang": "EN", "geo": "DE", "sinceTimePeriod": "2025-01"}, timeout=180)
print("status", r.status_code, len(r.content), flush=True)
if r.status_code != 200:
    print(r.text[:800])
    r = requests.get(B, params={"format": "JSON", "lang": "EN", "geo": "DE", "sinceTimePeriod": "2025-01"}, timeout=180)
    print("retry without unit:", r.status_code, len(r.content), r.text[:500])
j = r.json()
dims = j["id"]
print("dims:", dims, "sizes:", j["size"])
for d in dims:
    cat = j["dimension"][d]["category"]
    labels = cat.get("label", {})
    print(f"\n-- {d} ({len(labels)}):", json.dumps(labels)[:1800])
# sample values: for each nrg_bal, the 2025 months
idx = {d: j["dimension"][d]["category"]["index"] for d in dims}
size = j["size"]
import itertools
vals = j["value"]
def pos(coords):
    p = 0
    for d, s in zip(dims, size):
        p = p * s + coords[d]
    return p
bal = j["dimension"]["nrg_bal"]["category"]["index"]
time = j["dimension"]["time"]["category"]["index"]
print("\nGermany sample (GWh):")
for b, bi in bal.items():
    row = []
    for t, ti in time.items():
        c = {d: 0 for d in dims}
        c["nrg_bal"], c["time"] = bi, ti
        v = vals.get(str(pos(c)))
        row.append(None if v is None else round(v))
    print(b, j["dimension"]["nrg_bal"]["category"]["label"].get(b), row)
r2 = requests.get(B, params={"format": "JSON", "lang": "EN", "nrg_bal": "IC_CAL_MG", "sinceTimePeriod": "2025-06"}, timeout=180)
print("\nall geos, gross inland consumption:", r2.status_code)
if r2.status_code == 200:
    k = r2.json()
    print("geos:", list(k["dimension"]["geo"]["category"]["index"].keys()))
    print("times:", list(k["dimension"]["time"]["category"]["index"].keys()))
sys.exit(0)
