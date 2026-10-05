"""
Probe (benchmark only, never charted): Eurostat monthly gas imports (nrg_ti_gasm) and exports (nrg_te_gasm) by partner for DE, NL, BE, as TWh (GCV from TJ)
per month Oct 2022 - latest, to see in which months and on which partner the Germany + Netherlands balance (ENTSOG-based) departs from Eurostat. Prints
one line per geo/flow/partner (comma-separated monthly TWh), months listed first.
"""
import math
import sys

import requests

BASE = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"


def fetch(ds, geo):
    p = {"format": "JSON", "lang": "EN", "geo": geo, "siec": "G3000", "unit": "TJ_GCV", "sinceTimePeriod": "2022-10", "untilTimePeriod": "2026-12"}
    r = requests.get(BASE + ds, params=p, timeout=180)
    if r.status_code != 200:
        print(ds, geo, "HTTP", r.status_code, r.text[:100]); return {}, []
    j = r.json()
    dims, size = j["id"], j["size"]
    cat = {d: j["dimension"][d]["category"]["index"] for d in dims}
    inv = {d: {v: k for k, v in cat[d].items()} for d in dims}
    out = {}
    for k, v in j["value"].items():
        k = int(k); idx = {}
        for d, s in zip(reversed(dims), reversed(size)):
            idx[d] = inv[d][k % s]; k //= s
        out[(idx.get("partner"), idx["time"])] = out.get((idx.get("partner"), idx["time"]), 0.0) + v / 3600.0
    return out, sorted(cat["time"], key=lambda t: cat["time"][t])


for ds, label in (("nrg_ti_gasm", "IMP"), ("nrg_te_gasm", "EXP")):
    for geo in ("DE", "NL", "BE"):
        o, months = fetch(ds, geo)
        print(f"## {label} {geo} months {months[0] if months else ''}..{months[-1] if months else ''} n={len(months)}", flush=True)
        for pt in sorted({p for p, _ in o}):
            vals = [o.get((pt, m), 0.0) for m in months]
            if max(vals) >= 1.0:
                print(f"{label} {geo} {pt:6s} " + ",".join(f"{x:.1f}" for x in vals), flush=True)
sys.exit(0)
