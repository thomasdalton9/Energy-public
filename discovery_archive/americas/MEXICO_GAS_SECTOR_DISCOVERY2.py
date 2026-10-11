"""
Mexico gas demand by sector, round 2.  Round 1 (MEXICO_GAS_SECTOR_DISCOVERY.py) listed the open-data catalogue but did
not download the SENER 'Prontuario de Gas Natural y Petroquimicos' or the CENAGAS extractions file.  This one:

1. downloads those files and prints columns, row counts, date range and the distinct values of every low-cardinality
   text column (sector / use / user-type labels live there);
2. lists every dataset title and resource name for the demand / consumption / balance searches round 1 hid;
3. saves what it downloads (capped) under mx_raw/ for offline inspection.
Runs in GitHub Actions only (the editing sandbox blocks these hosts).
"""
import io
import os
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
BASE = "https://repodatos.atdt.gob.mx/api_update"
FILES = [
    f"{BASE}/sener/prontuario_gas_natural_petroquimicos/prontuario_datos_abiertos.csv",
    f"{BASE}/sener/prontuario_gas_natural_petroquimicos/prontuario_datos_abiertos_jul26.csv",
    f"{BASE}/cenagas/extracciones_inyecciones_gas_natural_sistrangas/CENEGAS_4_Capacidad_Historica_de_Extracciones_SISTRANGAS.csv",
    f"{BASE}/sener/informacion_transporte_almacenamiento_gas_natural_2026/bd_volumen_gas_natural_ta.csv",
    f"{BASE}/cne/gas_natural/gn_pre_hist_2017_2025.csv",
]
SEARCHES = ["demanda gas natural", "consumo gas natural", "balance nacional gas natural", "prospectiva gas natural",
            "ventas gas natural", "usuarios gas natural", "generacion electrica gas natural", "gas natural industria"]
CKAN = "https://www.datos.gob.mx/api/3/action/package_search"
MAX_SAVE = 8_000_000


def out(*a):
    print(*a, flush=True)


def read_any(raw):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, low_memory=False)
        except Exception:  # noqa: BLE001
            continue
    return None


os.makedirs("mx_raw", exist_ok=True)
for u in FILES:
    out(f"=========== {u}")
    try:
        r = requests.get(u, headers=H, timeout=(10, 180))
        out(f"HTTP {r.status_code} {len(r.content)}B last-modified={r.headers.get('Last-Modified')}")
        if r.status_code != 200:
            continue
        name = re.sub(r"[^A-Za-z0-9_.-]", "_", u.split("/")[-1])
        if len(r.content) <= MAX_SAVE:
            open(f"mx_raw/{name}", "wb").write(r.content)
        d = read_any(r.content)
        if d is None:
            out("  could not parse as CSV")
            continue
        out("rows", len(d), "columns", list(d.columns))
        out(d.head(4).to_string()[:1500])
        for c in d.columns:
            if d[c].dtype == object:
                n = d[c].nunique()
                if n <= 60:
                    out(f"  [{c}] {n} distinct: {sorted(map(str, d[c].dropna().unique()))[:60]}")
                elif re.search(r"fecha|date|mes|anio|ano|period", c, re.I):
                    out(f"  [{c}] {d[c].iloc[0]} .. {d[c].iloc[-1]}")
    except Exception as e:  # noqa: BLE001
        out("  failed", type(e).__name__, str(e)[:200])

out("=========== catalogue searches: every dataset and resource")
seen = set()
for q in SEARCHES:
    try:
        r = requests.get(CKAN, params={"q": q, "rows": 40}, headers=H, timeout=(10, 60))
        res = r.json()["result"]
        out(f"--- '{q}': {res['count']} datasets")
        for ds in res["results"]:
            if ds["id"] in seen:
                continue
            seen.add(ds["id"])
            out(f"  [{(ds.get('organization') or {}).get('title')}] {ds['title']}")
            for rs in ds.get("resources", [])[:6]:
                out(f"      {rs.get('format')} | {rs.get('name')} | {rs.get('url')}")
    except Exception as e:  # noqa: BLE001
        out(f"--- '{q}' failed {type(e).__name__}")
