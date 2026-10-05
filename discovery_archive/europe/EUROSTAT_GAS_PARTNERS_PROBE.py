"""
Probe (benchmark only, never charted): Eurostat monthly natural gas imports (nrg_ti_gasm) and exports (nrg_te_gasm) by partner country for the CEE /
Nordic / Baltic EU members, summed to annual TWh (GCV, from TJ) 2022-2025, to set against ENTSOG's border flows. Prints only.
"""
import sys
import requests

BASE = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
GEOS = "BG,EL,RO,HU,SK,CZ,PL,HR,SI,LT,LV,EE,FI,SE,AT".split(",")


def fetch(ds, geo):
    p = {"format": "JSON", "lang": "EN", "geo": geo, "siec": "G3000", "unit": "TJ_GCV", "sinceTimePeriod": "2022-01", "untilTimePeriod": "2025-12"}
    r = requests.get(BASE + ds, params=p, timeout=120)
    if r.status_code != 200:
        print(ds, geo, "HTTP", r.status_code, r.text[:100]); return {}
    j = r.json()
    dims, size = j["id"], j["size"]
    cat = {d: j["dimension"][d]["category"]["index"] for d in dims}
    inv = {d: {v: k for k, v in cat[d].items()} for d in dims}
    out = {}
    for k, v in j["value"].items():
        k = int(k); idx = {}
        for d, s in zip(reversed(dims), reversed(size)):
            idx[d] = inv[d][k % s]; k //= s
        key = (idx.get("partner"), idx["time"][:4])
        out[key] = out.get(key, 0.0) + v / 3600.0
    return out


for ds, label in (("nrg_ti_gasm", "IMPORTS"), ("nrg_te_gasm", "EXPORTS")):
    for geo in GEOS:
        o = fetch(ds, geo)
        parts = sorted({p for p, _ in o})
        print("=" * 60, label, geo, flush=True)
        for p in parts:
            vals = [o.get((p, str(y)), 0.0) for y in (2022, 2023, 2024, 2025)]
            if max(vals) >= 0.5:
                print(f"  {p:8s} " + " ".join(f"{x:7.1f}" for x in vals), flush=True)
sys.exit(0)
