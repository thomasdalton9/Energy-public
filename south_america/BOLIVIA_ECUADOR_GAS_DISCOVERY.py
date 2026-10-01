"""
Round 2 of the gas demand-by-use source search.

Bolivia: Ministerio de Hidrocarburos y Energias "Boletin Energetico
  Trimestral" (quarterly PDF, monthly rows?), its listing page, INE's
  "Boletin Sectorial de Hidrocarburos", YPFB's indicadores page. Looking
  for domestic-market gas by sector (termoelectrico, industrial, GNV,
  redes residencial/comercial, consumidores directos).
Ecuador: Petroecuador "Informe Estadistico" - find the listing of all
  monthly editions, and print the full DESPACHOS (volumes) page for the
  GAS NATURAL (pipeline gas to Termogas Machala) and GAS NATURAL
  LICUADO (Bajo Alto LNG for industry) rows, plus the Amistad
  production lines.
Not reachable from the editing sandbox.
"""
import io
import re
from urllib.parse import urljoin

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:30]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:120]}")
        return None


def links(r, pat, base):
    found = []
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I):
            found.append((urljoin(base, h), t))
    return found


def pdf_pages(content, page_pat, line_pat=None, max_lines=70, tables=True):
    if content[:4] != b"%PDF":
        out("  not a PDF:", content[:80])
        return
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not re.search(page_pat, text, re.I):
                continue
            out(f"  --- page {i + 1}")
            for ln in text.splitlines()[:max_lines]:
                if line_pat is None or re.search(line_pat, ln, re.I):
                    out("     ", ln[:200])
            if tables:
                for tb in page.extract_tables()[:2]:
                    out(f"     TABLE {len(tb)} rows")
                    for row in tb[:30]:
                        out("       ", [str(x)[:20].replace("\n", " ") if x else "" for x in row][:14])


# ------------------------------------------------------------------ Bolivia
out("=================== BOLIVIA: MHE Boletin Energetico Trimestral")
r = get("https://www.mhe.gob.bo/wp-content/uploads/2025/12/Boletin-trimestral-3T_2025-Intranet.pdf")
if r is not None and r.status_code == 200:
    pdf_pages(r.content, r"mercado interno|consumo.*gas|gas natural.*sector|termoel|GNV|redes de gas",
              max_lines=60)
for page in ("https://www.mhe.gob.bo/boletines/", "https://www.mhe.gob.bo/publicaciones/",
             "https://www.mhe.gob.bo/?s=boletin+energetico"):
    r = get(page)
    if r is not None and r.status_code == 200:
        for u, t in links(r, r"boletin|bolet[ií]n|\.pdf", page)[:60]:
            out(f"   {t[:80]!r} -> {u}")

out("\n=================== BOLIVIA: INE Boletin Sectorial de Hidrocarburos")
r = get("https://www.ine.gob.bo/index.php/boletin-sectorial-de-hidrocarburos-n-1-2025/")
if r is not None and r.status_code == 200:
    for u, t in links(r, r"\.pdf|\.xlsx?|descarg", "https://www.ine.gob.bo/")[:30]:
        out(f"   {t[:80]!r} -> {u}")
        if re.search(r"\.pdf", u, re.I):
            rr = get(u)
            if rr is not None and rr.status_code == 200:
                pdf_pages(rr.content, r"gas natural", r"gas|termo|industr|GNV|domest|resid|comerc|export|20\d\d",
                          tables=True)
            break
r = get("https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos/")
if r is not None and r.status_code == 200:
    for u, t in links(r, r"\.xlsx?|gas|hidrocarb", "https://www.ine.gob.bo/")[:40]:
        out(f"   {t[:80]!r} -> {u}")

out("\n=================== BOLIVIA: YPFB")
for page in ("https://www.ypfb.gob.bo/indicadores", "https://www.ypfb.gob.bo/es/informacion-institucional/boletines-estadisticos",
             "https://www.ypfb.gob.bo/boletines-estadisticos", "https://www.ypfb.gob.bo/publicaciones"):
    r = get(page)
    if r is not None and r.status_code == 200:
        txt = re.sub(r"<[^>]+>", " ", r.text)
        for m in re.finditer(r"(mercado interno|termoel|GNV|industrial|redes)[^.]{0,160}", txt, re.I):
            out("   txt:", re.sub(r"\s+", " ", m.group(0))[:180])
        for u, t in links(r, r"boletin|estad|\.pdf|\.xlsx?", page)[:40]:
            out(f"   {t[:80]!r} -> {u}")

# ------------------------------------------------------------------ Ecuador
out("\n=================== ECUADOR: Petroecuador listing")
for page in ("https://www.eppetroecuador.ec/?p=3721", "https://www.eppetroecuador.ec/?page_id=3721",
             "https://www.eppetroecuador.ec/?s=informe+estadistico"):
    r = get(page)
    if r is not None and r.status_code == 200:
        out("   title:", (re.search(r"<title>(.*?)</title>", r.text, re.S) or [None, ""])[1][:100])
        for u, t in links(r, r"download\.php|informe|estad", page)[:80]:
            out(f"   {t[:80]!r} -> {u}")

out("\n=================== ECUADOR: Informe Estadistico id=3089 - despachos and gas production")
r = get("https://www.eppetroecuador.ec/wp-content/plugins/download-monitor/download.php?id=3089")
if r is not None and r.status_code == 200:
    pdf_pages(r.content, r"DESPACHOS|Amistad|GAS NATURAL", r"DESPACHO|Cifras|Enero|GAS|Amistad|BLOQUE 6|pies|MMBTU|Producto",
              max_lines=200, tables=False)
