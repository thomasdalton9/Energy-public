"""
South & Southeast Asia hydro reservoir discovery, round 10 (after round 9: Internet Archive rate limits; Pakistan only): test the two new pulls end to end (twice each, the second
run exercises the incremental path) into temp workbooks, then build their proposed water-year charts with add_charts'
machinery and png_charts (registry entries injected here; add_charts.py itself is not changed). Also: the feed's own
'Total' row vs the six-reservoir sum for Sri Lanka.
"""
import os
import shutil
import subprocess
import sys
import tempfile

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "matplotlib"], check=False)


def out(*a):
    print(*a, flush=True)


import add_charts  # noqa: E402


def sri_lanka_reservoirs(p):
    d = add_charts._sheet(p, "Daily", "date")
    x = d["Total_storage_GWh"].dropna()
    med = x.rolling(15, center=True, min_periods=5).median()
    bad = (x / med - 1).abs() > 0.15   # one-day keying errors in the feed (e.g. a reservoir entered x3)
    out(f"  SL spikes dropped: {int(bad.sum())}: {x[bad].round(0).to_dict()}")
    x = x[~bad]
    return [{"name": "Storage", "water_year": x.resample("D").interpolate(limit=31, limit_area="inside"),
             "title": "Sri Lanka major hydro reservoir storage (PUCSL / CEB, six reservoirs)", "units": "GWh",
             "y_decimals": 0}]


def pakistan_reservoirs(p):
    d = add_charts._sheet(p, "Daily", "date")
    res = []
    for dam, what in (("Tarbela", "Indus"), ("Mangla", "Jhelum")):
        col = f"{dam}_level_ft"
        if col in d and d[col].notna().sum() >= 2:
            res.append({"name": dam, "water_year": d[col].dropna().resample("D").interpolate(limit=31, limit_area="inside"),
                        "title": f"{dam} reservoir water level ({what}; WAPDA / IRSA)", "units": "ft above sea level",
                        "y_decimals": 0, "sheet": f"Water year - {dam}"})
    return res


add_charts.REGISTRY["sri_lanka_hydro_reservoirs.xlsx"] = sri_lanka_reservoirs
add_charts.REGISTRY["pakistan_hydro_reservoirs.xlsx"] = pakistan_reservoirs
import png_charts  # noqa: E402

tmp = tempfile.mkdtemp()
for script, name in (("asia/PAKISTAN_IRSA_RESERVOIRS.py", "pakistan_hydro_reservoirs.xlsx"),):
    p = os.path.join(tmp, name)
    for run in (1, 2):
        out(f"\n--- {script} run {run}")
        res = subprocess.run([sys.executable, os.path.join(ROOT, script), "--out", p], capture_output=True, text=True,
                             timeout=1300, env={**os.environ, "WAYBACK_BUDGET": "1000"})
        out(res.stdout[-5000:])
        out(res.stderr[-3000:])
    if not os.path.exists(p):
        continue
    for sh, d in pd.read_excel(p, sheet_name=None, index_col=0).items():
        out(f"  [{sh}] {d.shape}")
        if sh in ("Daily", "Rainfall"):
            out(d.describe().T.round(1).to_string()[:3000])
            idx = pd.to_datetime(d.index)
            out("  days per year:", pd.Series(1, idx).groupby(idx.year).sum().to_dict())
            if "Source" in d:
                out("  by source:", d["Source"].value_counts().to_dict())
                w = d[d["Source"] == "IRSA (Wayback)"]
                out("  wayback days per month:", pd.Series(1, pd.to_datetime(w.index)).groupby(
                    pd.to_datetime(w.index).to_period("M")).sum().to_dict())
            out(d.tail(4).to_string()[:2000])
        elif sh == "Units":
            out("\n".join(map(str, d.index))[:3000])
        else:
            out(d.to_string()[:1500])
    n = add_charts.add_charts(p)
    out(f"  add_charts -> {n}")
    import openpyxl
    wb = openpyxl.load_workbook(p)
    out("  sheets:", wb.sheetnames)
    for ws in wb.worksheets:
        if ws._charts:
            c = ws._charts[0]
            out(f"   chart on {ws.title}: {type(c).__name__}, series {len(c.series)}")
    try:
        out("  png:", png_charts.render(p, out_dir=tmp))
    except Exception as e:  # noqa: BLE001
        out("  png err", e)

d = pd.read_excel(p, sheet_name="Daily", index_col=0)
print(d[d.Source == "IRSA (Wayback)"].iloc[:, [0, 1, 2, 3, 4, 5]].to_string())
