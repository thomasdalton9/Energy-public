"""
Chile reservoirs, third probe: DGA's daily-data front ends.
  - vipnet.mop.gob.cl (Visualizador Hidrometrico Nacional: reservoir volumes) and
    mapas2.mop.gob.cl (Estadistica Hidrometrica): save the pages and their JS
    bundles so the JSON endpoints behind them can be found;
  - snia.mop.gob.cl/BNAConsultas/reportes (official daily reports, JSF form) and
    the DGA SAT online map: save the pages;
  - DGA station lists for lakes/reservoirs (zips).
Everything goes to $PROBE_OUT (pushed to a scratch branch by the workflow).
"""
import os
import re
from urllib.parse import urljoin

import requests

OUT = os.environ.get("PROBE_OUT", "probe_out")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-CL,es;q=0.9"}
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers.update(H)


def name(u):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", re.sub(r"^https?://", "", u))[-150:]


def save(u, follow=False):
    try:
        r = S.get(u, timeout=(10, 120))
    except Exception as e:
        print(f"ERR {type(e).__name__} {u}: {str(e)[:150]}", flush=True)
        return None
    print(f"{r.status_code} {len(r.content):>9}B {r.headers.get('content-type', '')[:40]} {u} -> {r.url}", flush=True)
    if r.status_code == 200:
        with open(os.path.join(OUT, name(u)), "wb") as f:
            f.write(r.content)
        if follow and "html" in r.headers.get("content-type", ""):
            for m in re.finditer(r"""<script[^>]+src=["']([^"']+)["']""", r.text, re.I):
                js = urljoin(r.url, m.group(1))
                if re.search(r"googletag|jquery|bootstrap|analytics", js):
                    continue
                save(js)
            for m in re.finditer(r"""<link[^>]+href=["']([^"']+\.json)["']""", r.text, re.I):
                save(urljoin(r.url, m.group(1)))
    return r


for u in ["https://vipnet.mop.gob.cl/", "https://mapas2.mop.gob.cl/",
          "https://snia.mop.gob.cl/BNAConsultas/reportes",
          "https://snia.mop.gob.cl/sat/site/informes/mapas/mapas.xhtml",
          "https://snia.mop.gob.cl/dgasat/pages/dgasat_main/dgasat_main.htm"]:
    save(u, follow=True)
for u in ["https://dga.mop.gob.cl/uploads/sites/13/2024/07/EstacionesNivelesLagosEmbalses.zip",
          "https://dga.mop.gob.cl/uploads/sites/13/2024/07/Embalses-1.zip",
          "https://dga.mop.gob.cl/uploads/sites/13/2026/09/Informe-Nacional-25.09.2026.xlsx"]:
    save(u)
# common API guesses behind the visualiser
for u in ["https://vipnet.mop.gob.cl/api/", "https://vipnet.mop.gob.cl/api/embalses",
          "https://vipnet.mop.gob.cl/api/estaciones", "https://vipnet.mop.gob.cl/config.json",
          "https://vipnet.mop.gob.cl/assets/config.json"]:
    save(u)
