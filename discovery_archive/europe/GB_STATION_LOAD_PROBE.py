"""
Probe: can Great Britain's station load be derived from NESO historic demand (TSD - ND - pumping - exports)? Prints the NESO columns
and annual sums (TWh) of ND, TSD, embedded, pumping, interconnector imports/exports, for 2023-2025, plus the implied station load.
"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))) + "/europe")
import GB_POWER_DAILY as G  # noqa: E402

for y in (2023, 2024, 2025):
    rid = G.neso_resource(y)
    rows, off = [], 0
    while True:
        j = G.get(f"{G.NESO}/datastore_search", resource_id=rid, limit=32000, offset=off).json()["result"]
        rows += j["records"]
        off += len(j["records"])
        if not j["records"] or off >= j.get("total", 0):
            break
    d = pd.DataFrame(rows)
    if y == 2023:
        print("columns", d.columns.tolist(), flush=True)
    num = d.drop(columns=[c for c in ("SETTLEMENT_DATE", "SETTLEMENT_PERIOD", "_id") if c in d]).apply(pd.to_numeric, errors="coerce")
    flows = [c for c in num if c.endswith("_FLOW")]
    imp = num[flows].clip(lower=0).sum(axis=1)
    exp = (-num[flows].clip(upper=0)).sum(axis=1)
    out = pd.Series({"ND": num["ND"].sum() * .5e-6, "TSD": num["TSD"].sum() * .5e-6 if "TSD" in num else float("nan"),
                     "emb_wind": num["EMBEDDED_WIND_GENERATION"].sum() * .5e-6, "emb_solar": num["EMBEDDED_SOLAR_GENERATION"].sum() * .5e-6,
                     "pump": num["PUMP_STORAGE_PUMPING"].sum() * .5e-6, "imports": imp.sum() * .5e-6, "exports": exp.sum() * .5e-6})
    out["station_load_implied"] = out["TSD"] - out["ND"] - out["pump"] - out["exports"]
    print(y, out.round(2).to_dict(), flush=True)
sys.exit(0)
