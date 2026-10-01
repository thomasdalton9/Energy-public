"""
Monthly gas demand-by-use sources for Chile and Ecuador (public PDFs /
open data), to see whether a Colombia-style monthly split can be built.

Chile:   CNE "Reporte Mensual Sector Energetico"
         https://www.cne.cl/wp-content/uploads/YYYY/MM/RMensual_vYYYYMM.pdf
Ecuador: Petroecuador "Informe Estadistico" (monthly PDF, download.php?id=N),
         datosabiertos.gob.ec CKAN datasets, and the central bank's quarterly
         "Boletin Analitico del Sector Petrolero".
Prints every page/line/table mentioning natural gas so a parser can be
written. Not reachable from the editing sandbox.
"""
import io
import re

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
GAS = re.compile(r"gas natural|\bGNL\b|\bGN\b|\bLNG\b|regasific|Quintero|Mejillones|Amistad|Bajo Alto|Termogas|MMpc|MPCD|pies c", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:120]}")
        return None


def gas_pages(content, max_tables=3, max_lines=60):
    if content[:4] != b"%PDF":
        out("  not a PDF:", content[:100])
        return
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages; first page: {(pdf.pages[0].extract_text() or '')[:200]!r}")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            hits = [ln for ln in text.splitlines() if GAS.search(ln)]
            if not hits:
                continue
            out(f"  --- page {i + 1} ({len(hits)} gas lines)")
            for ln in text.splitlines()[:max_lines]:
                out("     ", ln[:200])
            for tb in page.extract_tables()[:max_tables]:
                out(f"     TABLE {len(tb)} rows")
                for row in tb[:25]:
                    out("       ", [str(x)[:22].replace("\n", " ") if x else "" for x in row][:14])


# ------------------------------------------------------------------ Chile
out("=================== CHILE: CNE Reporte Mensual")
latest = None
for y, m in [(2026, mm) for mm in range(10, 0, -1)]:
    u = f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/RMensual_v{y}{m:02d}.pdf"
    r = get(u)
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        latest = r
        break
if latest is not None:
    gas_pages(latest.content)
out("\n-- CNE 2021 report URL pattern probe")
for y, m in [(2021, 1), (2021, 6), (2022, 1), (2023, 1), (2024, 1)]:
    get(f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/RMensual_v{y}{m:02d}.pdf")
r = get("https://www.cne.cl/nuestros-servicios/reportes/informacion-y-estadisticas/")
if r is not None:
    for h in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
        if re.search(r"gas|estad|reporte|mensual|xlsx|csv", h, re.I):
            out("   link:", h[:160])

# ------------------------------------------------------------------ Ecuador
out("\n=================== ECUADOR: Petroecuador Informe Estadistico")
for i in (3089, 3068):
    r = get(f"https://www.eppetroecuador.ec/wp-content/plugins/download-monitor/download.php?id={i}")
    if r is not None and r.status_code == 200:
        gas_pages(r.content, max_tables=2, max_lines=45)
r = get("https://www.eppetroecuador.ec/?p=3721")  # transparency / statistics page guess
for page in ("https://www.eppetroecuador.ec/?page_id=22735", "https://www.eppetroecuador.ec/informes-estadisticos/"):
    r = get(page)
    if r is not None and r.status_code == 200:
        for h, t in re.findall(r'href="([^"]*download[^"]*)"[^>]*>([^<]{0,100})<', r.text):
            out(f"   {t.strip()[:80]!r} -> {h}")

out("\n=================== ECUADOR: datosabiertos.gob.ec")
for q in ("gas natural", "produccion mensual petroecuador", "hidrocarburos gas"):
    r = get("https://datosabiertos.gob.ec/api/3/action/package_search", params={"q": q, "rows": 15})
    if r is None or r.status_code != 200:
        continue
    try:
        for p in r.json()["result"]["results"]:
            out(f"  [{q}] {p['name']}: {p.get('title', '')[:90]}")
            for res in p.get("resources", [])[:6]:
                out(f"       {res.get('format')} {res.get('name', '')[:60]!r} {res.get('url')}")
    except Exception as e:
        out("  parse error", e)
r = get("https://datosabiertos.gob.ec/api/3/action/package_show", params={"id": "produccion-mensual-petroecuador"})
if r is not None and r.status_code == 200:
    for res in r.json()["result"].get("resources", []):
        out(f"   resource {res.get('format')} {res.get('name')} {res.get('url')}")
        if str(res.get("format", "")).lower() in ("csv", "xlsx", "xls"):
            rr = get(res["url"])
            if rr is not None and rr.status_code == 200:
                txt = rr.content[:3000]
                out("     head:", txt[:1500])

out("\n=================== ECUADOR: BCE quarterly oil-sector bulletin")
r = get("https://contenido.bce.fin.ec/documentos/Estadisticas/Hidrocarburos/ASP202502.pdf")
if r is not None and r.status_code == 200:
    gas_pages(r.content, max_tables=2, max_lines=40)
