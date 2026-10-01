"""
One-off probe: find a public daily series of Argentina's hydro reservoir
levels (cota, m a.s.l.) and stored volume (hm3 / % useful volume) for the
Comahue lakes (El Chocon, Piedra del Aguila, Alicura, Cerros Colorados /
Los Barreales / Mari Menuco, Pichi Picun Leufu), Futaleufu, and flows /
levels at Yacyreta and Salto Grande - ideally with 10+ years of history.

ARGENTINA_CAMMESA_DISCOVERY.py (Sep-2026) already found that CAMMESA's
PROGRAMACION_DIARIA and PARTE_POST_OPERATIVO carry no lake levels, and
that AIC (aic.gob.ar/embalses) has one 'Ver Detalle' page per lake.

Round 1 here checks, printing what each returns:
  A. AIC: the lake list + detail pages, scripts/iframes/JSON endpoints
     behind them, and any history/download links.
  B. BDHI / SNIH (snih.hidricosargentina.gob.ar): station search pages.
  C. CAMMESA: cammesaweb pages that mention hidro/embalse/cota (and the
     document 'nemo' codes they use), then those nemo codes on pub-svc.
  D. datos.gob.ar / datos.energia.gob.ar CKAN search + series API search.
  E. INA alerta, EBY (Yacyreta), CTM Salto Grande home pages.
Run by discovery_archive/workflows/argentina_hydro_discovery.yml.
"""

print("STARTING", flush=True)

import datetime as dt
import json
import re
import sys
from urllib.parse import urljoin

import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
KEY = re.compile(r"cota|embals|volumen|hm3|hm³|caudal|nivel|chocon|chocón|piedra del|alicur|barreales|mari menuco|"
                 r"pichi|futaleuf|yacyret|salto grande|hidrol|hist[oó]ric|descarg|xls|csv|json|api", re.I)
ROUND = sys.argv[1] if len(sys.argv) > 1 else "1"
S = requests.Session()
S.headers.update(UA)


def text_of(html):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", html, flags=re.S | re.I))


def get(url, show_text=1500, links=True, max_links=60, **kw):
    try:
        r = S.get(url, timeout=kw.pop("timeout", 45), **kw)
    except requests.RequestException as e:
        print(f"\n==== {url}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None
    ct = r.headers.get("Content-Type", "")
    print(f"\n==== {url}: {r.status_code} {ct} {len(r.content):,} bytes (final {r.url})", flush=True)
    if "html" in ct or "text" in ct or "json" in ct or "javascript" in ct:
        body = r.text
        if "json" in ct:
            print("  JSON", body[:show_text], flush=True)
            return r
        t = text_of(body)
        m = re.search(r"<title>(.*?)</title>", body, re.S | re.I)
        print(f"  TITLE {m.group(1).strip()[:120] if m else ''}", flush=True)
        if show_text:
            print(f"  TEXT {t[:show_text]}", flush=True)
        if links:
            n = 0
            for href, label in re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', body, re.S | re.I):
                label = text_of(label).strip()
                full = urljoin(r.url, href)
                if KEY.search(full + " " + label) and n < max_links:
                    print(f"  LINK {label[:70]!r} -> {full}", flush=True)
                    n += 1
            for src in re.findall(r'<(?:script|iframe)[^>]+src=["\']([^"\']+)', body, re.I)[:25]:
                print(f"  SRC {urljoin(r.url, src)}", flush=True)
            for u in sorted(set(re.findall(r'["\']((?:https?:)?//[^"\'\s]*|/[^"\'\s]*)(?:\.json|/api/|ajax|\.php\?|\.aspx|\.ashx|\.xlsx?|\.csv)[^"\'\s]*["\']', body)))[:20]:
                print(f"  DATAURL? {u}", flush=True)
            for f in re.findall(r"<form[^>]*>", body, re.I)[:5]:
                print(f"  FORM {f[:200]}", flush=True)
    return r


# ---------------------------------------------------------------- A. AIC
def aic():
    print("\n######## A. AIC", flush=True)
    r = get("https://www.aic.gob.ar/embalses", show_text=2500)
    if r is None:
        return
    details = sorted(set(re.findall(r'href=["\']([^"\']*embalses-detalle\?a=\d+[^"\'#]*)', r.text)))
    print(f"\n#### {len(details)} AIC reservoir detail pages: {details}", flush=True)
    for k, href in enumerate(details):
        d = get(urljoin(r.url, href), show_text=2500 if k < 2 else 900, links=k < 2)
        if d is not None and k == 0:
            # inline scripts may call a JSON endpoint for the chart / history
            for s in re.findall(r"<script[^>]*>(.*?)</script>", d.text, re.S | re.I):
                if re.search(r"ajax|fetch|url|json|data|chart|serie", s, re.I) and len(s) < 20000:
                    flat = re.sub(r"\s+", " ", s)
                    print(f"  INLINE SCRIPT {flat[:1500]}", flush=True)
    for u in ["https://www.aic.gob.ar/", "https://www.aic.gob.ar/sitio/hidrologia", "https://www.aic.gob.ar/hidrologia",
              "https://www.aic.gob.ar/informes", "https://www.aic.gob.ar/sitio/informes",
              "https://www.aic.gob.ar/estado-de-embalses", "https://www.aic.gob.ar/datos-hidrometeorologicos"]:
        get(u, show_text=600)


# ---------------------------------------------------------------- B. BDHI / SNIH
def bdhi():
    print("\n######## B. BDHI / SNIH", flush=True)
    for u in ["https://snih.hidricosargentina.gob.ar/", "https://snih.hidricosargentina.gob.ar/Filtros.aspx",
              "https://snih.hidricosargentina.gob.ar/MuestraDatos.aspx", "http://bdhi.hidricosargentina.gob.ar/",
              "https://www.argentina.gob.ar/obras-publicas/hidricas/base-de-datos-hidrologica-integrada"]:
        get(u, show_text=1500)


# ---------------------------------------------------------------- C. CAMMESA
LOOKUP_URL = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
ATTACHMENT_URL = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
TIME_FMT = "%Y-%m-%dT%H:%M:%S.000Z"


def nemo_docs(nemo, days=10):
    end = dt.date.today()
    params = {"fechadesde": (end - dt.timedelta(days=days)).strftime(TIME_FMT),
              "fechahasta": (end + dt.timedelta(days=1)).strftime(TIME_FMT), "nemo": nemo}
    try:
        r = requests.get(LOOKUP_URL, params=params, headers={"User-Agent": "gas-demand-scripts/1.0"}, timeout=45)
    except requests.RequestException as e:
        print(f"  nemo {nemo}: FAILED {e}", flush=True)
        return []
    try:
        docs = r.json()
    except ValueError:
        docs = None
    n = len(docs) if isinstance(docs, list) else 0
    print(f"  nemo {nemo}: {r.status_code}, {n} docs", flush=True)
    if isinstance(docs, list):
        for d in docs[:3]:
            print(f"    doc id={d.get('id')} titulo={d.get('titulo')!r} fecha={d.get('fecha')} "
                  f"adjuntos={[a.get('id') for a in d.get('adjuntos', [])][:10]}", flush=True)
        return docs
    return []


def cammesa():
    print("\n######## C. CAMMESA", flush=True)
    nemos = set()
    pages = []
    for sm in ["https://cammesaweb.cammesa.com/sitemap.xml", "https://cammesaweb.cammesa.com/wp-sitemap.xml",
               "https://cammesaweb.cammesa.com/sitemap_index.xml", "https://cammesaweb.cammesa.com/page-sitemap.xml"]:
        r = get(sm, show_text=0, links=False)
        if r is None or not r.ok:
            continue
        locs = re.findall(r"<loc>(.*?)</loc>", r.text)
        print(f"  {len(locs)} locs", flush=True)
        for loc in locs:
            if loc.endswith(".xml") and loc not in (sm,):
                rr = get(loc, show_text=0, links=False)
                if rr is not None and rr.ok:
                    pages += re.findall(r"<loc>(.*?)</loc>", rr.text)
            else:
                pages.append(loc)
    pages = sorted(set(pages))
    print(f"\n  {len(pages)} cammesaweb pages", flush=True)
    hits = [p for p in pages if re.search(r"hidr|embals|cota|caudal|hidraul|sintesis|post-oper|parte|semanal", p, re.I)]
    for p in pages:
        print(f"   PAGE {p}", flush=True)
    for p in hits[:25]:
        r = get(p, show_text=500)
        if r is not None:
            found = set(re.findall(r"nemo[\"'=:\s]+[\"']?([A-Z][A-Z0-9_]{3,})", r.text))
            print(f"  NEMOS on page: {sorted(found)}", flush=True)
            nemos |= found
    guesses = ["DATOS_HIDRAULICOS", "DATOS_HIDROLOGICOS", "HIDROLOGIA", "INFORME_HIDROLOGICO", "PARTE_HIDRAULICO",
               "COTAS", "EMBALSES", "COTAS_EMBALSES", "INFORME_SINTESIS", "INFORME_SINTESIS_MENSUAL",
               "SINTESIS_MENSUAL", "INFORME_MENSUAL", "PROGRAMACION_SEMANAL", "PROG_SEMANAL", "PROGRAMACION_ESTACIONAL",
               "INFORME_SEMANAL", "PARTE_SEMANAL", "HIDRAULICIDAD", "PARTE_POST_OPERATIVO", "INFORME_HIDRAULICO",
               "HIDRO", "CAUDALES", "SITUACION_HIDROLOGICA", "ESTADO_EMBALSES"]
    print("\n  -- nemo checks", flush=True)
    for n in sorted(nemos | set(guesses)):
        nemo_docs(n)


# ---------------------------------------------------------------- D. open data
def open_data():
    print("\n######## D. open data", flush=True)
    for base in ["https://datos.gob.ar", "http://datos.energia.gob.ar", "https://datos.energia.gob.ar"]:
        for q in ["embalse", "embalses", "cota", "hidroelectric", "caudal"]:
            u = f"{base}/api/3/action/package_search?q={q}&rows=20"
            try:
                r = S.get(u, timeout=45)
                j = r.json()
            except Exception as e:  # noqa: BLE001
                print(f"  {u}: FAILED {type(e).__name__}: {str(e)[:150]}", flush=True)
                continue
            res = j.get("result", {})
            print(f"\n  {u}: {res.get('count')} packages", flush=True)
            for p in res.get("results", [])[:20]:
                print(f"    PKG {p.get('name')}: {p.get('title')}", flush=True)
                for rs in p.get("resources", [])[:6]:
                    print(f"       RES {rs.get('format')} {rs.get('name')!r} {rs.get('url')}", flush=True)
    for q in ["embalse", "cota", "caudal", "chocon", "hidraulica"]:
        get(f"https://apis.datos.gob.ar/series/api/search/?q={q}&limit=20", show_text=2500)


# ---------------------------------------------------------------- E. others
def others():
    print("\n######## E. INA / EBY / CTM", flush=True)
    for u in ["https://www.ina.gob.ar/alerta/", "https://alerta.ina.gob.ar/pub/gui", "https://alerta.ina.gob.ar/a5/",
              "https://www.eby.org.ar/", "https://www.eby.org.ar/index.php/hidrologia",
              "https://www.saltogrande.org/", "https://www.saltogrande.org/hidrologia",
              "https://www.argentina.gob.ar/orsep"]:
        get(u, show_text=700)


if ROUND == "1":
    for fn in (aic, bdhi, cammesa, open_data, others):
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            print(f"!! {fn.__name__} crashed: {type(e).__name__}: {e}", flush=True)

print("\nDONE", flush=True)
