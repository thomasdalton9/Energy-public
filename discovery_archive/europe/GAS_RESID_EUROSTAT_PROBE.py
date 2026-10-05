"""Probe (benchmark only): Eurostat nrg_cb_gasm - every balance item for ES, IT, FR, UK (siec G3000, TJ GCV -> TWh), annual 2023-2025 and monthly 2025-2026 - and
nrg_ti_gasm / nrg_te_gasm imports/exports by partner for ES, IT, FR (annual 2022-2025, monthly 2025). Saves results/residual/eurostat_*.csv."""
import os
import requests, pandas as pd
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
BASE = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
def fetch(ds, geo, extra=None, since="2022-01"):
    p = {"format": "JSON", "lang": "EN", "geo": geo, "siec": "G3000", "unit": "TJ_GCV", "sinceTimePeriod": since, "untilTimePeriod": "2026-12"}
    p.update(extra or {})
    r = requests.get(BASE + ds, params=p, timeout=180)
    if r.status_code != 200:
        print(ds, geo, r.status_code, r.text[:120]); return pd.DataFrame()
    j = r.json(); dims, size = j["id"], j["size"]
    cat = {d: j["dimension"][d]["category"]["index"] for d in dims}
    inv = {d: {v: k for k, v in cat[d].items()} for d in dims}
    rec = []
    for k, v in j["value"].items():
        k = int(k); idx = {}
        for d, s in zip(reversed(dims), reversed(size)):
            idx[d] = inv[d][k % s]; k //= s
        idx["TWh"] = v / 3600.0; rec.append(idx)
    return pd.DataFrame(rec)
bal, par = [], []
for geo in ["ES", "IT", "FR", "UK"]:
    d = fetch("nrg_cb_gasm", geo); d["geo"] = geo; bal.append(d); print("bal", geo, len(d), flush=True)
    for ds in ("nrg_ti_gasm", "nrg_te_gasm"):
        if geo == "UK": continue
        d = fetch(ds, geo); d["geo"] = geo; d["ds"] = ds; par.append(d); print(ds, geo, len(d), flush=True)
pd.concat(bal).to_csv(os.path.join(OUT, "eurostat_balance_items.csv"), index=False)
pd.concat(par).to_csv(os.path.join(OUT, "eurostat_partners.csv"), index=False)
