"""
Discovery round 3 (after CAPACITY_PBUE_DISCOVERY2.py).

Round 2 found:
  Peru: COES pages list their folders through a JS file browser (hidden hfBaseDirectory, e.g.
        'Operación/Estudios/Potencia Efectiva/'); POST /Portal/browser/vistadatos with guessed fields -> 500.
        Annual statistics page POSTs Publicaciones/Estadisticas/Detalle {anio}.
  Bolivia: cndc.bo routes include historico/potencia (desde, hasta), historico/potencia/detalle (anio),
        dashboard/potencia, dashboard/potencia-historial (modo, anio, mes), ga/potencia-efectiva, oferta-potencia.
  Uruguay: MIEM 'Potencia instalada por central.xlsx' -> sheet 'Pot Inst Fuente' (annual, MW, since 1967).
  Ecuador: controlrecursosyenergia.gob.ec now serves a CentOS-WebPanel test page (also over plain HTTP) -
        the ARCERNNR site itself is down; CENACE Informe Anual posts exist for 2017-2022 only.

This round: COES browser JS + listing; CNDC capacity endpoints; Ecuador alternative hosts.
"""
import json
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get, links  # noqa: E402


def snippet(text, pat, width=300, limit=20):
    for i, m in enumerate(re.finditer(pat, text, re.I)):
        if i >= limit:
            break
        print("     ~", re.sub(r"\s+", " ", text[max(0, m.start() - width):m.end() + width]))


def peru():
    print("\n================ PERU")
    r = get("https://www.coes.org.pe/Portal/Operacion/Estudios/PotenciaEfectiva")
    if r is None:
        return
    scripts = re.findall(r'<script[^>]+src="([^"]+)"', r.text)
    print("     scripts:", scripts)
    snippet(r.text, r"hfBaseDirectory|browserDocument|controlador|siteRoot", width=400, limit=10)
    for s in scripts:
        if re.search(r"browser|operacion|estudio|potencia|coes", s, re.I) and "jquery" not in s.lower():
            js = get("https://www.coes.org.pe" + s if s.startswith("/") else s)
            if js is not None and js.status_code == 200:
                snippet(js.text, r"\$\.ajax|url\s*:|vistadatos|download", width=250, limit=25)
    for js_url in ["https://www.coes.org.pe/Portal/Content/Scripts/browser.js",
                   "https://www.coes.org.pe/Portal/Content/Scripts/browser/browser.js",
                   "https://www.coes.org.pe/Portal/Areas/Operacion/Content/Scripts/estudios.js"]:
        js = get(js_url)
        if js is not None and js.status_code == 200 and "javascript" in js.headers.get("content-type", ""):
            snippet(js.text, r"\$\.ajax|url\s*:", width=250, limit=25)
    r = get("https://www.coes.org.pe/Portal/Publicaciones/Estadisticas/Detalle", method="POST", data={"anio": 2025})
    if r is not None and r.status_code == 200:
        for href, text in re.findall(r'href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S)[:200]:
            t = re.sub(r"<[^>]+>|\s+", " ", text).strip()
            print(f"     detalle: {t[:90]!r} -> {href[:250]}")
        snippet(r.text, r"potencia|capacidad", width=200, limit=15)


def bolivia():
    print("\n================ BOLIVIA")
    base = "https://www.cndc.bo/wp-json/cndc/v1/"
    for path, params in [("historico/potencia", {"desde": 2015, "hasta": 2026}),
                         ("historico/potencia/detalle", {"anio": 2025}),
                         ("historico/potencia/detalle", {"anio": 2021}),
                         ("dashboard/potencia", {}),
                         ("dashboard/potencia-historial", {"modo": "mensual", "anio": 2025, "mes": 8}),
                         ("dashboard/potencia-historial", {"modo": "anual", "anio": 2025}),
                         ("dashboard/potencia-historial", {}),
                         ("ga/potencia-efectiva", {}),
                         ("oferta-potencia", {"fecha": "2026-09-15"}),
                         ("oferta-potencia", {"fecha": "2021-01-15"}),
                         ("estadisticas/categorias", {}),
                         ("estadisticas/categorias", {"tipo": "anual"}),
                         ("memorias", {})]:
        r = get(base + path, params=params)
        if r is not None and r.status_code == 200:
            print("     ", r.text[:2500].replace("\n", " "))


def ecuador():
    print("\n================ ECUADOR")
    for u in ["https://www.arconel.gob.ec/", "https://arconel.gob.ec/", "https://www.regulacionelectrica.gob.ec/",
              "https://www.arcernnr.gob.ec/", "https://www.energia.gob.ec/", "https://www.recursosyenergia.gob.ec/",
              "https://www.ambienteyenergia.gob.ec/", "https://www.cenace.gob.ec/informe-anual-2022/",
              "https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm"]:
        print(f"\n -- {u}")
        links(u, pat=re.compile(r"estad|potencia|capacidad|anual|informe|boletin|balance|\.xlsx?|\.pdf", re.I),
              limit=40)
    r = get("https://www.cenace.gob.ec/wp-json/wp/v2/posts", params={"search": "informe anual", "per_page": 20,
                                                                    "_fields": "id,date,title,link"})
    if r is not None and r.status_code == 200:
        for p in r.json():
            print(f"     post {p.get('date')} {p.get('title', {}).get('rendered')} {p.get('link')}")
    r = get("https://www.cenace.gob.ec/wp-json/wp/v2/pages", params={"per_page": 100, "_fields": "id,title,link"})
    if r is not None and r.status_code == 200:
        for p in r.json():
            print(f"     page {p.get('title', {}).get('rendered')} {p.get('link')}")


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "bolivia", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
