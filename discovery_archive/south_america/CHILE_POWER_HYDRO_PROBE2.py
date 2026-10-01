"""
Second Chile probe (after CHILE_POWER_HYDRO_DISCOVERY.py found coordinador.cl
returns 403 to GitHub runners, CNE publishes Generacion_Bruta.xlsx via its
WordPress media library and DGA publishes weekly 'Boletin hidrometeorologico'
PDFs):
  cen  - why the 403 (headers/body), other UAs, other CEN hosts
  cne  - Generacion_Bruta.xlsx layout; the CNE electricity statistics page links
  ea   - Energia Abierta electricity catalogue links
  dga  - DGA station statistics/visualiser pages, latest weekly bulletin's reservoir table
"""
import io
import re
import sys
from urllib.parse import urljoin

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA, "Accept-Language": "es-CL,es;q=0.9",
     "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"}
T = (10, 90)


def out(*a):
    print(*a, flush=True)


def get(url, headers=H, **kw):
    try:
        r = requests.get(url, headers=headers, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:140]}")
        return None


def page_links(url, pat, maxn=80):
    r = get(url)
    if r is None or r.status_code != 200:
        return []
    found = []
    for m in re.finditer(r"""(?:href|src|data-src|data-url)\s*=\s*["']([^"'#]+)["']""", r.text, re.I):
        u = urljoin(r.url, m.group(1).strip())
        if re.search(pat, u, re.I) and u not in found:
            found.append(u)
    for m in re.finditer(r"""["'](https?://[^"'\s<>]+)["']""", r.text):
        if re.search(pat, m.group(1), re.I) and m.group(1) not in found:
            found.append(m.group(1))
    for u in found[:maxn]:
        out("    ", u)
    return found


def show_xlsx(content, rows=14, tail=4):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    for ws in wb.worksheets:
        data = [r for r in ws.iter_rows(values_only=True)]
        out(f"  SHEET {ws.title!r}: {len(data)} rows x {max((len(r) for r in data), default=0)} cols")
        for r in data[:rows]:
            out("     ", [str(c)[:18] for c in r[:16] if c is not None])
        out("      ...")
        for r in data[-tail:]:
            out("     ", [str(c)[:18] for c in r[:16] if c is not None])


section = sys.argv[1] if len(sys.argv) > 1 else "all"

if section in ("all", "cen"):
    out("=========== CEN")
    r = get("https://www.coordinador.cl/")
    if r is not None:
        out("  headers:", {k: v for k, v in r.headers.items() if k.lower() in
                           ("server", "cf-ray", "x-cache", "via", "x-iinfo", "x-cdn", "set-cookie", "x-sucuri-id")})
        out("  body:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:400])
    for ua in ["curl/8.5.0", "python-requests/2.32", "Googlebot/2.1 (+http://www.google.com/bot.html)"]:
        get("https://www.coordinador.cl/reportes-y-estadisticas/", headers={"User-Agent": ua})
    for u in ["https://coordinador.cl/", "http://www.coordinador.cl/",
              "https://www.coordinador.cl/wp-content/uploads/2024/01/",
              "https://infotecnica.coordinador.cl/", "https://api-infotecnica.coordinador.cl/v1/centrales/?format=json",
              "https://infotecnica.coordinador.cl/api/v1/centrales/", "https://reportes.coordinador.cl/",
              "https://portal.api.coordinador.cl/", "https://www.cen.cl/", "https://sic.coordinador.cl/",
              "https://datos.coordinador.cl/", "https://transparencia.coordinador.cl/",
              "https://www.coordinador.cl/feed/", "https://web.archive.org/web/2026/https://www.coordinador.cl/reportes-y-estadisticas/"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            out("   ", re.sub(r"\s+", " ", r.text[:300]))

if section in ("all", "cne"):
    out("=========== CNE")
    for u in ["https://www.cne.cl/wp-content/uploads/2026/09/Generacion_Bruta.xlsx",
              "https://www.cne.cl/wp-content/uploads/2025/03/Generacion_Bruta_SSMM.xlsx"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            show_xlsx(r.content)
    for u in ["https://www.cne.cl/estadisticas/electricidad/", "https://www.cne.cl/normativas/electrica/estadisticas/",
              "https://www.cne.cl/estadisticas/", "https://www.cne.cl/estadisticas/energia/electricidad/"]:
        page_links(u, r"\.(xlsx?|csv|zip)|generaci|embalse|cota")
    r = get("https://www.cne.cl/wp-json/wp/v2/pages?search=estad&per_page=50&_fields=link,title")
    if r is not None and r.status_code == 200:
        for it in r.json():
            out("    ", it["link"], "|", it["title"]["rendered"][:60])
    for t in ["horaria", "Generacion_Bruta", "cota", "volumen", "hidro", "SEN"]:
        r = get(f"https://www.cne.cl/wp-json/wp/v2/media?search={t}&per_page=40&_fields=date,source_url")
        if r is not None and r.status_code == 200:
            for it in r.json():
                if re.search(r"\.(xlsx?|csv|zip|rar)$", it["source_url"], re.I):
                    out("    ", it["date"][:10], it["source_url"])

if section in ("all", "ea"):
    out("=========== Energia Abierta")
    page_links("https://energiaabierta.cl/categorias-estadistica/electricidad/", r"generaci|embalse|cota|hidro|dataset|api|\.(csv|xlsx)", 120)
    for u in ["https://energiaabierta.cl/?lang=&s=generacion&t=datasets-estadistica",
              "https://energiaabierta.cl/?lang=&s=embalse&t=datasets-estadistica",
              "https://energiaabierta.cl/?lang=&s=cota&t=datasets-estadistica",
              "http://desarrolladores.energiaabierta.cl/", "https://api.energiaabierta.cl/"]:
        page_links(u, r"generaci|embalse|cota|visualizaciones|dataset|api|\.(csv|xlsx)", 60)

if section in ("all", "dga"):
    out("=========== DGA")
    for u in ["https://dga.mop.gob.cl/estadisticas-estaciones-dga/", "https://dga.mop.gob.cl/visualizadores/",
              "https://dga.mop.gob.cl/sistema-hidrometrico-en-linea/", "https://dga.mop.gob.cl/servicios-de-informacion/boletines/",
              "https://dga.mop.gob.cl/boletines-hidrologicos-2/"]:
        page_links(u, r"embalse|\.(xlsx?|csv|zip)|snia|dgasat|estadistic|visualiz|boletin|hidrometr|arcgis|powerbi", 60)
    import pdfplumber
    for u in ["https://dga.mop.gob.cl/uploads/sites/13/2026/01/2026-09-28_Boletin-hidrometeorologico.pdf"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            out(f"  {len(pdf.pages)} pages")
            for i, p in enumerate(pdf.pages):
                txt = p.extract_text() or ""
                if re.search(r"embalse", txt, re.I):
                    out(f"  --- page {i + 1}")
                    for ln in txt.splitlines()[:70]:
                        out("     ", ln[:160])
out("DONE")
