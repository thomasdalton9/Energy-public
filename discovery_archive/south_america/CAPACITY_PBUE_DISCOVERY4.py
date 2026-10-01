"""
Discovery round 4 (after CAPACITY_PBUE_DISCOVERY3.py).

Round 3 found:
  Peru: COES annual statistics (Publicaciones/Estadisticas/Detalle {anio}) link per-chapter Excel files,
        e.g. 'Capitulo 02_ESTADO DE LA INFRAESTRUCTURA DEL SEIN.xlsx', 2017-2025. The page file browser is
        driven by /Portal/Content/Scripts/web.js.
  Bolivia: cndc.bo /dashboard/potencia-historial?modo=mensual&anio=&mes= -> 12 months of capacity (MW) by
        technology ending at anio/mes; /historico/potencia/detalle?anio= -> capacity per plant by technology.
  Ecuador: ARCERNNR's statistics now live at https://arconel.gob.ec/estadistica-del-sector-electrico/ (valid TLS).

This round: COES chapter-2 workbook + web.js browser calls + boletines; CNDC history depth; ARCONEL statistics.
"""
import io
import re
import sys
from urllib.parse import quote

import pandas as pd

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get, links  # noqa: E402
from CAPACITY_PBUE_DISCOVERY2 import dump_excel  # noqa: E402
from CAPACITY_PBUE_DISCOVERY3 import snippet  # noqa: E402

COES = "https://www.coes.org.pe/Portal/"


def peru():
    print("\n================ PERU")
    js = get(COES + "Content/Scripts/web.js?v=8.5")
    if js is not None and js.status_code == 200:
        snippet(js.text, r"browser|vistadatos|\$\.ajax", width=300, limit=30)
    for year in (2025, 2021):
        url = COES + "browser/download?url=" + quote(f"Publicaciones/Estadisticas Anuales/{year}/Excel/"
                                                     "Capítulo 02_ESTADO DE LA INFRAESTRUCTURA DEL SEIN.xlsx")
        r = get(url)
        if r is not None and r.status_code == 200:
            dump_excel(r.content, f"cap02 {year}", rows=60)
    for u in [COES + "Publicaciones/Boletines/index?path=" + quote("Publicaciones/Boletines/"),
              COES + "Publicaciones/Estadisticas/index?path=" + quote("Publicaciones/Estadisticas Anuales/2025/")]:
        r = get(u)
        if r is not None and r.status_code == 200:
            snippet(r.text, r"\.xlsx|\.pdf|\.zip|hfBaseDirectory", width=200, limit=20)


def bolivia():
    print("\n================ BOLIVIA")
    base = "https://www.cndc.bo/wp-json/cndc/v1/"
    for anio, mes in [(2026, 9), (2026, 12), (2021, 12), (2020, 12), (2019, 12)]:
        r = get(base + "dashboard/potencia-historial", params={"modo": "mensual", "anio": anio, "mes": mes})
        if r is not None and r.status_code == 200:
            print("     ", r.text[:1200])
    for anio in (2026, 2024, 2022):
        r = get(base + "historico/potencia/detalle", params={"anio": anio})
        if r is not None and r.status_code == 200:
            for g in r.json().get("grupos", []):
                print(f"     {anio} {g.get('tec')}: " + "; ".join(f"{p['central']}={p['mw']}" for p in g["plantas"]))
    r = get(base + "estadisticas/documentos", params={"categoria_id": 235, "agrupado": "true"})
    if r is not None and r.status_code == 200:
        print("      diesel docs:", r.text[:800])


def ecuador():
    print("\n================ ECUADOR")
    found = links("https://arconel.gob.ec/estadistica-del-sector-electrico/", limit=120)
    for url, text in found:
        if re.search(r"\.xlsx?$", url, re.I) and re.search(r"potencia|capacidad|instalad|efectiv", url + text, re.I):
            r = get(url)
            if r is not None and r.status_code == 200:
                dump_excel(r.content, url, rows=40)
            break
    sub = [u for u, t in found if "arconel.gob.ec" in u and not re.search(r"\.(pdf|xlsx?|zip)$", u, re.I)]
    for u in sub[:12]:
        print(f"\n -- {u}")
        links(u, pat=re.compile(r"potencia|capacidad|instalad|efectiv|estad|boletin|anual|mensual|\.xlsx?|\.pdf",
                                re.I), limit=40)


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "bolivia", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
