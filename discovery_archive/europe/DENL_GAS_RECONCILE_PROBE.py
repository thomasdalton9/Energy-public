"""
Probe: independent figures to reconcile the Germany + Netherlands gas balance (ENTSOG-based, europe_master) for 2023-2024. Prints only.
  1. Eurostat nrg_ti_gasm (imports by partner, monthly) and nrg_te_gasm (exports by partner) for DE, NL, BE: TWh per year by partner and fuel.
  2. Eurostat nrg_cb_gasm balance items for DE, NL, BE (supply, imports, exports, stock change, consumption), TWh per year.
  3. CBS StatLine 86103NED (aardgasbalans, monthly): every column, annual sums.
"""
import math
import sys

import pandas as pd
import requests

EU = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
pd.set_option("display.width", 250, "display.max_rows", 500, "display.max_columns", 50)


def eurostat(ds, **flt):
    p = [("format", "JSON"), ("lang", "EN"), ("sinceTimePeriod", "2023-01"), ("untilTimePeriod", "2024-12")] + [(k, v) for k, vs in flt.items() for v in (vs if isinstance(vs, list) else [vs])]
    r = requests.get(EU + ds, params=p, timeout=300)
    print(ds, r.status_code, len(r.content), flush=True)
    if not r.ok:
        print(r.text[:300]); return pd.DataFrame()
    j = r.json()
    dims, size = j["id"], j["size"]
    idx = {d: list(j["dimension"][d]["category"]["index"]) for d in dims}
    st = [math.prod(size[i + 1:]) for i in range(len(size))]
    rows = []
    for k, v in j["value"].items():
        k = int(k)
        c = {d: idx[d][(k // st[i]) % size[i]] for i, d in enumerate(dims)}
        c["OBS_VALUE"] = float(v)
        rows.append(c)
    d = pd.DataFrame(rows)
    return d.rename(columns={"time": "TIME_PERIOD"})


for ds in ("nrg_ti_gasm", "nrg_te_gasm"):
    d = eurostat(ds, geo=["DE", "NL", "BE"], unit="TJ_GCV")
    if d.empty:
        continue
    print(d.columns.tolist())
    d["TWh"] = d["OBS_VALUE"] / 3600.0
    d["year"] = d["TIME_PERIOD"].str[:4]
    g = d.groupby(["geo", "siec", "partner", "year"])["TWh"].sum().unstack("year").round(1)
    g = g[(g.abs() > 0.5).any(axis=1)]
    print(f"== {ds} TWh per year (GCV) ==")
    print(g.to_string(), flush=True)

d = eurostat("nrg_cb_gasm", geo=["DE", "NL", "BE"], unit="TJ_GCV", siec="G3000")
if len(d):
    d["TWh"] = d["OBS_VALUE"] / 3600.0
    d["year"] = d["TIME_PERIOD"].str[:4]
    print("== nrg_cb_gasm balance items TWh per year ==")
    print(d.groupby(["geo", "nrg_bal", "year"])["TWh"].sum().unstack("year").round(1).to_string(), flush=True)

CBS = "https://opendata.cbs.nl/ODataApi/odata/86103NED/"
try:
    cols = requests.get(CBS + "DataProperties?$format=json", timeout=120).json()["value"]
    keys = {c["Key"]: c.get("Title") for c in cols if c.get("Type") in ("Dimension", "TimeDimension", "Topic", "GeoDimension")}
    print({k: v for k, v in keys.items()})
    rows, url = [], CBS + "TypedDataSet?$format=json"
    while url:
        j = requests.get(url, timeout=300).json()
        rows += j["value"]
        url = j.get("odata.nextLink")
    df = pd.DataFrame(rows)
    print(df.shape, df.columns.tolist())
    per = next((c for c in df.columns if "Perioden" in c), None)
    df["year"] = df[per].astype(str).str[:4]
    num = [c for c in df.columns if df[c].dtype != object and c not in ("ID",)]
    print("== CBS 86103NED annual sums (units as published, see titles) ==")
    out = df[df["year"].isin(["2023", "2024"]) & df[per].astype(str).str.contains("MM")].groupby("year")[num].sum().T
    out.index = [f"{c} | {keys.get(c, '')}" for c in out.index]
    print(out.round(1).to_string())
except Exception as e:  # noqa: BLE001
    print("CBS failed", type(e).__name__, e)
sys.exit(0)
