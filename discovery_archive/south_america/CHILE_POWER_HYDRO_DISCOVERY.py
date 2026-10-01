"""
Chile: key-free sources for daily generation by technology and hydro
reservoir levels/volumes (the CEN SIP API needs CEN_USER_KEY, so it is
not used). Probes:
  - CEN coordinador.cl pages + its WordPress REST API (media search for
    uploaded Excel/CSV/ZIP files), public JSON behind the charts;
  - CNE cne.cl and Energia Abierta (energiaabierta.cl, datos.energiaabierta.cl);
  - DGA (dga.mop.gob.cl, snia.mop.gob.cl) reservoir bulletins.
Prints statuses and every interesting link. Run in GitHub Actions only.
"""
import json
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-CL,es;q=0.9,en;q=0.8"}
T = (10, 60)
KEY = re.compile(r"generaci|embalse|cota|hidro|estad|reporte|operaci|tecnolog|energia|volum|boletin|novedad|"
                 r"\.xlsx?\b|\.csv\b|\.zip\b|\.json\b|api|dataset|descarga", re.I)
FILE = re.compile(r"\.(xlsx?|csv|zip|json|pdf)(\?|$)", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:50]} {len(r.content)}B final={r.url}")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:160]}")
        return None


def links(url, maxn=150, only=KEY):
    r = get(url)
    if r is None or r.status_code != 200 or "html" not in r.headers.get("content-type", ""):
        if r is not None and r.status_code == 200:
            out("   body:", r.text[:300].replace("\n", " "))
        return []
    html = r.text
    found = []
    for m in re.finditer(r"""(?:href|src|data-src|data-url|action)\s*=\s*["']([^"'#]+)["']""", html, re.I):
        u = urljoin(r.url, m.group(1).strip())
        if only.search(u) and u not in found:
            found.append(u)
    # bare URLs inside scripts (e.g. JSON/endpoints)
    for m in re.finditer(r"""["'](https?://[^"'\s<>]+)["']""", html):
        u = m.group(1)
        if (FILE.search(u) or re.search(r"api|json|ajax|datos|powerbi|tableau|embed", u, re.I)) and u not in found:
            found.append(u)
    out(f"   {len(found)} links (title: {re.search(r'<title>(.*?)</title>', html, re.S).group(1).strip()[:80] if '<title>' in html else ''})")
    for u in found[:maxn]:
        out("    ", u)
    return found


def wp_media(site, terms):
    for t in terms:
        url = f"{site}/wp-json/wp/v2/media?search={t}&per_page=100&orderby=date&order=desc"
        r = get(url)
        if r is None or r.status_code != 200:
            continue
        try:
            items = r.json()
        except Exception:
            out("   not json:", r.text[:200])
            continue
        out(f"   total={r.headers.get('X-WP-Total')} pages={r.headers.get('X-WP-TotalPages')}")
        for it in items[:100]:
            out("    ", it.get("date", "")[:10], it.get("source_url"))


def wp_pages(site, terms, kinds=("pages", "posts")):
    for k in kinds:
        for t in terms:
            r = get(f"{site}/wp-json/wp/v2/{k}?search={t}&per_page=50&_fields=link,title,date")
            if r is None or r.status_code != 200:
                continue
            try:
                for it in r.json():
                    out("    ", it.get("date", "")[:10], it.get("link"), "|", (it.get("title") or {}).get("rendered", "")[:80])
            except Exception:
                out("   not json:", r.text[:200])


section = sys.argv[1] if len(sys.argv) > 1 else "all"

if section in ("all", "cen"):
    out("=================== CEN coordinador.cl")
    for u in ["https://www.coordinador.cl/",
              "https://www.coordinador.cl/reportes-y-estadisticas/",
              "https://www.coordinador.cl/operacion/",
              "https://www.coordinador.cl/operacion/graficos/operacion-real/",
              "https://www.coordinador.cl/operacion/graficos/operacion-real/generacion-real/",
              "https://www.coordinador.cl/operacion/documentos/",
              "https://www.coordinador.cl/operacion/documentos/operacion-real/",
              "https://www.coordinador.cl/operacion/graficos/operacion-real/cotas-y-volumenes-de-embalses/",
              "https://www.coordinador.cl/operacion/graficos/operacion-real/embalses/",
              ]:
        links(u)
    out("--- WP REST pages/posts")
    wp_pages("https://www.coordinador.cl", ["generacion", "embalse", "cota", "estadistica", "reporte diario", "novedades"])
    out("--- WP REST media")
    wp_media("https://www.coordinador.cl", ["generacion", "embalse", "cota", "estadistica", "reporte", "tecnologia",
                                            "energia", "novedades", "operacion"])

if section in ("all", "cne"):
    out("=================== CNE + Energia Abierta")
    for u in ["https://www.cne.cl/", "https://www.cne.cl/estadisticas/", "https://www.cne.cl/estadisticas/electricidad/",
              "https://www.cne.cl/normativas/electrica/", "https://www.cne.cl/nuestros-servicios/reportes/",
              "https://energiaabierta.cl/", "https://energiaabierta.cl/electricidad/",
              "https://energiaabierta.cl/?s=generacion", "https://energiaabierta.cl/?s=embalse",
              "http://datos.energiaabierta.cl/", "https://datos.energiaabierta.cl/home/",
              "https://datos.energiaabierta.cl/search/?q=generacion",
              ]:
        links(u)
    out("--- WP REST media cne.cl / energiaabierta.cl")
    wp_media("https://www.cne.cl", ["generacion", "embalse", "electricidad", "capacidad"])
    wp_media("https://energiaabierta.cl", ["generacion", "embalse", "cota"])
    wp_pages("https://energiaabierta.cl", ["generacion", "embalse"])

if section in ("all", "dga"):
    out("=================== DGA")
    for u in ["https://dga.mop.gob.cl/", "https://www.dga.cl/",
              "https://dga.mop.gob.cl/servicios-de-informacion/",
              "https://dga.mop.gob.cl/informacion-hidrometeorologica/",
              "https://dga.mop.gob.cl/estadisticas-hidrometeorologicas/",
              "https://dga.mop.gob.cl/boletines/",
              "https://snia.mop.gob.cl/", "https://snia.mop.gob.cl/BNAConsultas/reportes",
              "https://snia.mop.gob.cl/dgasat/", "https://snia.mop.gob.cl/sat/site/informes/mapas/mapas.xhtml",
              ]:
        links(u)
    wp_media("https://dga.mop.gob.cl", ["embalse", "boletin", "hidrometeorologic"])
    wp_pages("https://dga.mop.gob.cl", ["embalse", "boletin"])
out("DONE")
