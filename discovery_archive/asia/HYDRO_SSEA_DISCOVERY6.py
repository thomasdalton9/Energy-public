"""
South & Southeast Asia hydro reservoir discovery, round 6 (after HYDRO_SSEA_DISCOVERY5.py):
  Vietnam   VNDMS layer list (LayerData/GetLayerDisplay) and the EVN hydro-lake layer (/evnlake?...): fields per lake,
            any per-lake history endpoint
  Test      run asia/SRI_LANKA_PUCSL_RESERVOIRS.py and asia/PAKISTAN_IRSA_RESERVOIRS.py to temp workbooks (twice, to
            check the incremental path) and print summaries
"""
import html
import json
import os
import re
import subprocess
import sys
import tempfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "X-Requested-With": "XMLHttpRequest", "Referer": "https://vndms.gov.vn/"}


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    try:
        r = requests.get(u, headers=H, timeout=(15, 90), **k)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {e}")
        return None


def strip(x):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " | ", x or ""))).strip()


def vndms():
    out("\n=========== VNDMS layers")
    r = get("https://vndms.gov.vn/LayerData/GetLayerDisplay")
    if r is not None and r.ok:
        def walk(x, depth=0):
            if isinstance(x, dict):
                if "layer_name" in x:
                    out("  " * depth + f"- {x.get('id')} {x.get('layer_name')!r} type {x.get('layer_type')} url {x.get('service_url')!r}")
                for v in x.values():
                    if isinstance(v, (list, dict)):
                        walk(v, depth + 1)
            elif isinstance(x, list):
                for v in x:
                    walk(v, depth)
        walk(r.json())
    for q in ("evnlake?typelake=true", "evnlake?dangxa=0&typelake=true", "evnlake", "evnlake?typelake=false",
              "evnlake?vuotmucdang=0&typelake=true", "evnlake?dangxa=1&typelake=true"):
        r = get("https://vndms.gov.vn/" + q)
        if r is None or not r.ok:
            continue
        try:
            fs = r.json().get("features", [])
        except ValueError:
            out("  not json", r.text[:200])
            continue
        out(f"  {q}: {len(fs)} features")
        for f in fs[:3]:
            p = dict(f.get("properties", {}))
            pi = p.pop("popupInfo", "")
            out("    props:", json.dumps(p, ensure_ascii=False)[:800])
            out("    popup:", strip(pi)[:1500])
            out("    popup raw:", pi[:1500])
        if fs:
            out("    labels:", [f["properties"].get("label") for f in fs][:200])
    r = get("https://vndms.gov.vn/bundles/vndms-components-ca5dfe7d0b.min.js")
    if r is not None:
        for m in list(re.finditer(r"evnlake|EvnLake|evn_lake|lake", r.text))[:20]:
            out("    js: " + r.text[max(0, m.start() - 300):m.start() + 300].replace("\n", " "))
    r = get("https://vndms.gov.vn/bundles/templates/modals/station-modal.tmpl")
    if r is not None:
        out("  station-modal:", r.text[:4000])


def test_pulls():
    out("\n=========== test pulls")
    tmp = tempfile.mkdtemp()
    for script, name in (("asia/SRI_LANKA_PUCSL_RESERVOIRS.py", "sl.xlsx"), ("asia/PAKISTAN_IRSA_RESERVOIRS.py", "pk.xlsx")):
        p = os.path.join(tmp, name)
        for run in (1, 2):
            out(f"--- {script} run {run}")
            res = subprocess.run([sys.executable, script, "--out", p], capture_output=True, text=True, timeout=900)
            out(res.stdout[-6000:])
            out(res.stderr[-3000:])
        if os.path.exists(p):
            for sh, d in pd.read_excel(p, sheet_name=None, index_col=0).items():
                out(f"  [{sh}] {d.shape}")
                if sh in ("Daily", "Rainfall"):
                    out(d.describe().T.to_string()[:3000])
                    out(d.head(3).to_string()[:1500])
                    out(d.tail(3).to_string()[:1500])
                    idx = pd.to_datetime(d.index)
                    out("  days per year:", pd.Series(1, idx).groupby(idx.year).sum().to_dict())
                else:
                    out(d.to_string()[:2000])


for f in (vndms, test_pulls):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
