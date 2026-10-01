"""
Downloads candidate Chile files/pages into $PROBE_OUT so they can be
inspected offline (the discovery workflow pushes them to a scratch branch).
Usage: CHILE_POWER_HYDRO_FETCH.py <set> [extra URLs...]
"""
import json
import os
import re
import sys

import requests

OUT = os.environ.get("PROBE_OUT", "probe_out")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-CL,es;q=0.9"}

SETS = {
    "first": [
        "https://www.cne.cl/wp-content/uploads/2026/09/Generacion_Bruta.xlsx",
        "https://energiaabierta.cl/categorias-estadistica/electricidad/?_sft_etiquetas-estadistica=hidrologia",
        "https://energiaabierta.cl/categorias-estadistica/electricidad/?_sft_etiquetas-estadistica=generacion-bruta",
        "https://energiaabierta.cl/categorias-estadistica/electricidad/?_sft_etiquetas-estadistica=generacion",
        "https://energiaabierta.cl/categorias-estadistica/electricidad/",
        "https://dga.mop.gob.cl/uploads/sites/13/2026/01/2026-09-28_Boletin-hidrometeorologico.pdf",
        "https://dga.mop.gob.cl/uploads/sites/13/2025/01/Boletin-Hidrometrico-DGA-septiembre-2025.pdf",
        "https://dga.mop.gob.cl/estadisticas-estaciones-dga/",
        "https://dga.mop.gob.cl/visualizadores/",
        "https://dga.mop.gob.cl/sistema-hidrometrico-en-linea/",
        "https://dga.mop.gob.cl/servicios-de-informacion/boletines/",
        "https://web.archive.org/web/2026/https://www.coordinador.cl/reportes-y-estadisticas/",
        "https://web.archive.org/web/2026/https://www.coordinador.cl/operacion/graficos/operacion-real/",
        "https://api-infotecnica.coordinador.cl/v1/",
        "https://api-infotecnica.coordinador.cl/swagger/",
    ],
    "urls": [],
}
WPMEDIA = {"first": [("https://dga.mop.gob.cl", "Boletin"), ("https://dga.mop.gob.cl", "embalse"),
                     ("https://dga.mop.gob.cl", "Informe-Hidrometeorologico")]}


def name(u):
    n = re.sub(r"^https?://", "", u)
    return re.sub(r"[^A-Za-z0-9._-]+", "_", n)[-150:]


s = sys.argv[1] if len(sys.argv) > 1 else "first"
os.makedirs(OUT, exist_ok=True)
for u in SETS.get(s, []) + sys.argv[2:]:
    try:
        r = requests.get(u, headers=H, timeout=(10, 120))
        print(f"{r.status_code} {len(r.content):>9}B {r.headers.get('content-type', '')[:40]} {u}", flush=True)
        if r.status_code == 200:
            with open(os.path.join(OUT, name(u)), "wb") as f:
                f.write(r.content)
    except Exception as e:
        print(f"ERR {type(e).__name__} {u}: {str(e)[:120]}", flush=True)
for site, term in WPMEDIA.get(s, []):
    items, page = [], 1
    while True:
        r = requests.get(f"{site}/wp-json/wp/v2/media", params={"search": term, "per_page": 100, "page": page,
                         "_fields": "date,source_url,title"}, headers=H, timeout=(10, 120))
        if r.status_code != 200:
            break
        items += r.json()
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            break
        page += 1
    print(f"wp media {site} {term}: {len(items)}", flush=True)
    with open(os.path.join(OUT, name(f"{site}_media_{term}") + ".json"), "w") as f:
        json.dump(items, f, indent=0)
