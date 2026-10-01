"""
Discovery round 2 (after CAPACITY_PBUE_DISCOVERY.py).

Round 1 found:
  Peru: COES 'Potencia Efectiva' page (/Portal/Operacion/Estudios/PotenciaEfectiva), 'Boletines Mensuales'
        (/Portal/Publicaciones/Boletines/), annual statistics (/Portal/publicaciones/estadisticas/estadistica?anio=).
        minem.gob.pe times out.
  Bolivia: cndc.bo wp-json route index is a list (parse fixed here).
  Uruguay: MIEM 'Series estadisticas de energia electrica' has 'Potencia instalada por central (.zip)';
        ADME annual report data files (Informe_anual_potencias_2024.xlsx ...).
  Ecuador: controlrecursosyenergia.gob.ec serves a certificate for webserver.arconel.gob.ec (hostname
        mismatch, not a missing intermediate); cenace.gob.ec works with its AIA intermediate.
  Ember yearly CSV has Capacity (GW) by fuel to 2025 for all four.

This round dumps the actual content of those leads.
"""
import io
import json
import re
import sys
import zipfile
from urllib.parse import urljoin

import pandas as pd

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get, links  # noqa: E402

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_colwidth", 40)


def show_html(url, pat=r"browser|vistadatos|\.xlsx|\.pdf|\.zip|ajax|url\s*:|data\s*:", width=160, limit=40):
    r = get(url)
    if r is None or r.status_code != 200:
        return None
    hits = 0
    for m in re.finditer(pat, r.text, re.I):
        print("     ~", re.sub(r"\s+", " ", r.text[max(0, m.start() - width):m.end() + width]))
        hits += 1
        if hits >= limit:
            break
    return r


def dump_excel(content, name="", rows=25):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print(f"     not excel: {e}")
        return
    print(f"     {name} sheets: {xl.sheet_names}")
    for s in xl.sheet_names[:12]:
        d = pd.read_excel(xl, sheet_name=s, header=None)
        print(f"     --- sheet {s!r} shape {d.shape}")
        print(d.head(rows).to_string(max_colwidth=28)[:4000])


def peru():
    print("\n================ PERU")
    show_html("https://www.coes.org.pe/Portal/Operacion/Estudios/PotenciaEfectiva")
    print()
    show_html("https://www.coes.org.pe/Portal/Publicaciones/Boletines/")
    print()
    show_html("https://www.coes.org.pe/Portal/publicaciones/estadisticas/estadistica?anio=2025")
    # COES file browser (the pages above list folders through it)
    for folder in ["Operación/Estudios/Potencia Efectiva", "Publicaciones/Boletines", "Publicaciones/Estadisticas",
                   "Publicaciones/Estadisticas Anuales", "Post Operación/Informes/Evaluación Mensual"]:
        for ep in ["https://www.coes.org.pe/Portal/browser/vistadatos",
                   "https://www.coes.org.pe/Portal/Browser/Vistadatos"]:
            r = get(ep, method="POST", data={"baseDirectory": folder, "url": folder, "indicador": "S"})
            if r is not None and r.status_code == 200:
                print("     ", re.sub(r"\s+", " ", r.text[:1500]))
                break


def bolivia():
    print("\n================ BOLIVIA")
    r = get("https://www.cndc.bo/wp-json/cndc/v1")
    if r is not None and r.status_code == 200:
        j = r.json()
        routes = j.get("routes", {})
        for route, info in routes.items():
            eps = info.get("endpoints") or []
            args = eps[0].get("args", {}) if eps and isinstance(eps[0], dict) else {}
            print(f"     route {route} {info.get('methods')} args={list(args) if isinstance(args, dict) else args}")
    r = get("https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos", params={"categoria_id": 155,
                                                                                  "agrupado": "true"})
    if r is not None and r.status_code == 200:
        j = r.json()
        grupos = j.get("grupos", [])
        print("     keys:", list(j), "groups:", len(grupos))
        if grupos:
            print("     group0:", json.dumps(grupos[0], ensure_ascii=False)[:1500])
        types = {}
        for g in grupos:
            for d in g.get("documentos", []):
                types.setdefault(d.get("tipo_documento") or d.get("nombre") or d.get("titulo"), []).append(
                    (g.get("periodo") or g.get("label") or g.get("nombre"), d.get("archivo_url")))
        for t, v in types.items():
            print(f"     type {t!r}: {len(v)}; first {v[0]}; last {v[-1]}")


def uruguay():
    print("\n================ URUGUAY")
    page = "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/series-estadisticas-energia-electrica"
    found = links(page, pat=re.compile(r"\.zip|\.xlsx?", re.I))
    for url, text in found:
        if re.search(r"potencia", text + url, re.I):
            r = get(url)
            if r is None or r.status_code != 200:
                continue
            try:
                z = zipfile.ZipFile(io.BytesIO(r.content))
            except zipfile.BadZipFile:
                dump_excel(r.content, url)
                continue
            for n in z.namelist():
                print(f"     zip member {n} {z.getinfo(n).file_size}B")
                if n.lower().endswith((".xlsx", ".xls", ".csv")):
                    data = z.read(n)
                    if n.lower().endswith(".csv"):
                        print(data[:3000].decode("latin-1"))
                    else:
                        dump_excel(data, n, rows=40)
    for u in ["https://adme.com.uy/db-docs/Docs_secciones/nid_526/Informe_anual_potencias_2024.xlsx"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            dump_excel(r.content, u)


def ecuador():
    print("\n================ ECUADOR")
    for u in ["https://webserver.arconel.gob.ec/", "https://www.controlrecursosyenergia.gob.ec/",
              "http://www.controlrecursosyenergia.gob.ec/"]:
        r = get(u, allow_redirects=False)
        if r is not None:
            print("      location:", r.headers.get("location"), re.sub(r"\s+", " ", r.text[:300]))
    for q in ["informe anual", "potencia efectiva", "capacidad instalada", "estadistica"]:
        for typ in ["search"]:
            r = get(f"https://www.cenace.gob.ec/wp-json/wp/v2/{typ}", params={"search": q, "per_page": 30})
            if r is not None and r.status_code == 200:
                try:
                    for it in r.json():
                        print(f"     [{q}] {it.get('title')} -> {it.get('url')} ({it.get('subtype')})")
                except ValueError:
                    print(r.text[:300])
    r = get("https://www.cenace.gob.ec/wp-json/wp/v2/dlm_download", params={"search": "anual", "per_page": 50})
    if r is not None and r.status_code == 200:
        try:
            for it in r.json():
                print(f"     dlm {it.get('id')} {it.get('title', {}).get('rendered')} {it.get('link')}")
        except ValueError:
            print(r.text[:300])
    links("https://www.cenace.gob.ec/informe-anual-2017/", pat=re.compile(r"informe|anual", re.I))
    # Wayback copy of ARCERNNR's statistics page (normal TLS on web.archive.org)
    for u in ["https://web.archive.org/web/2026/https://www.controlrecursosyenergia.gob.ec/estadistica-del-sector-electrico/",
              "https://web.archive.org/web/2025/https://www.controlrecursosyenergia.gob.ec/"]:
        links(u, limit=60)


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "bolivia", "uruguay", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
