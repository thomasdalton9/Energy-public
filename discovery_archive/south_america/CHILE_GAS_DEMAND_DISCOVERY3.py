"""
Chile gas demand by sector, round 3. Round 2 (CHILE_GAS_DEMAND_DISCOVERY2.py) found on
CNE Estadisticas > Hidrocarburo:
  - Consumo_mensual_GN-1.xls (2015 upload): monthly gas consumption by sector (Mm3:
    Res-Com, Industrial, GNC, Generacion, Petroquimica y Refineria, Otros) from 2006
  - Produccion_combustibles-junio-2024.xlsx: monthly oil + gas production, ENAP and
    CEOP (private operators), thousand m3, from 2018 (source: Ministerio de Energia)
This round prints each file's full date range and last rows, looks for newer
versions of both, greps the CNE Anuario Estadistico de Energia 2021 for gas
consumption by sector, lists the INE "Produccion de electricidad, gas y agua" files
and ENAP's integrated reports page. Runs in GitHub Actions.
"""
import io
import re

import pandas as pd
import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)


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


def links(r, pat, base=""):
    if r is None or r.status_code != 200:
        return []
    res, seen = [], set()
    for h, t in re.findall(r'href="([^"]+)"[^>]*>(.{0,200}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if h.startswith("/"):
            h = base + h
        if h in seen or not re.search(pat, h + " " + t, re.I):
            continue
        seen.add(h)
        res.append((h, t))
        out(f"   link: {t[:80]!r} -> {h[:220]}")
    return res


def tail_sheet(content, sheet=None, head=8, tail=40):
    x = pd.ExcelFile(io.BytesIO(content))
    out("  sheets:", x.sheet_names)
    for s in ([sheet] if sheet else x.sheet_names):
        df = pd.read_excel(x, sheet_name=s, header=None)
        out(f"  [{s}] shape {df.shape}")
        for part in (df.head(head), df.tail(tail)):
            for i, row in part.iterrows():
                vals = [str(v)[:20] for v in row.tolist()[:16] if str(v) != "nan"]
                if vals:
                    out(f"     {i}: {vals}")
            out("     ...")


# ------------------------------------------------------------------ 1) consumption by sector file
out("=================== Consumo_mensual_GN")
for u in ["https://www.cne.cl/wp-content/uploads/2015/05/Consumo_mensual_GN-1.xls",
          "https://www.cne.cl/wp-content/uploads/2015/05/Consumo_mensual_GN.xls"]:
    r = get(u)
    if r is not None and r.status_code == 200 and len(r.content) > 1000:
        tail_sheet(r.content, "Consumo_GN", head=6, tail=25)

# ------------------------------------------------------------------ 2) production file + newer versions
out("\n=================== Produccion_combustibles")
r = get("https://www.cne.cl/wp-content/uploads/2024/09/Produccion_combustibles-junio-2024.xlsx")
if r is not None and r.status_code == 200:
    tail_sheet(r.content, "produccion crudo gas m_energia", head=6, tail=45)
    tail_sheet(r.content, "Pet. Crudo - GN - Carbon", head=0, tail=15)
for q in ["Produccion_combustibles", "Produccion combustibles", "Consumo_mensual", "consumo mensual GN",
          "Venta_mensual", "importaciones-web"]:
    rr = get("https://www.cne.cl/wp-json/wp/v2/media", params={"search": q, "per_page": 100})
    if rr is not None and rr.status_code == 200:
        for s in re.findall(r'"source_url":"([^"]+)"', rr.text):
            out(f"   [{q}] {s.replace(chr(92) + '/', '/')}")
# the current importaciones-web workbook (may hold gas by origin/terminal)
r = get("https://www.cne.cl/wp-content/uploads/2026/06/importaciones-web.xlsx")
if r is not None and r.status_code == 200:
    x = pd.ExcelFile(io.BytesIO(r.content))
    out("  importaciones-web sheets:", x.sheet_names)
    for s in x.sheet_names[:4]:
        df = pd.read_excel(x, sheet_name=s, header=None)
        out(f"  [{s}] shape {df.shape}")
        for i, row in pd.concat([df.head(8), df.tail(6)]).iterrows():
            vals = [str(v)[:16] for v in row.tolist()[:16] if str(v) != "nan"]
            if vals:
                out(f"     {i}: {vals}")

# ------------------------------------------------------------------ 3) Anuario 2021: gas tables
out("\n=================== CNE Anuario Estadistico de Energia 2021")
r = get("https://www.cne.cl/wp-content/uploads/2022/07/AnuarioEstadisticoEnergia2021.pdf")
if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        shown = 0
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if re.search(r"gas natural", t, re.I) and re.search(r"consumo|sector|industri|residencial|generaci",
                                                                   t, re.I):
                shown += 1
                if shown > 10:
                    break
                out(f"  ---- page {i + 1}")
                for ln in t.splitlines()[:70]:
                    out("     ", ln[:200])

# ------------------------------------------------------------------ 4) INE production of electricity, gas and water
out("\n=================== INE produccion de electricidad, gas y agua")
for u in ["https://www.ine.gob.cl/estadisticas-por-tema/industria-energia-y-construccion/produccion-de-electricidad-gas-y-agua",
          "https://www.ine.gob.cl/estadisticas-por-tema/industria-energia-y-construccion/estructura-de-la-electricidad-gas-y-agua"]:
    r = get(u)
    ls = links(r, r"\.xlsx?|\.csv|docs/default-source|boletin|cuadro", "https://www.ine.gob.cl")
    for h, t in ls:
        if re.search(r"\.xlsx?($|\?)", h, re.I):
            rr = get(h.replace("&amp;", "&"))
            if rr is not None and rr.status_code == 200:
                try:
                    tail_sheet(rr.content, None, head=12, tail=6)
                except Exception as e:
                    out("  read error", e)
            break

# ------------------------------------------------------------------ 5) ENAP integrated reports
out("\n=================== ENAP reportes integrados / memorias")
r = get("https://www.enap.cl/reporte-integrado-y-memorias")
ls = links(r, r"memoria|reporte|integrado|files/get|\.pdf", "https://www.enap.cl")
cand = [h for h, t in ls if re.search(r"2025|2024", t) and re.search(r"files/get|\.pdf", h)]
if cand:
    rr = get(cand[0])
    if rr is not None and rr.status_code == 200 and rr.content[:4] == b"%PDF":
        with pdfplumber.open(io.BytesIO(rr.content)) as pdf:
            out(f"  {len(pdf.pages)} pages")
            n = 0
            for i, p in enumerate(pdf.pages):
                for ln in (p.extract_text() or "").splitlines():
                    if re.search(r"gas natural|MMm3|millones de m|m3/d|Magallanes", ln, re.I) and re.search(r"\d", ln):
                        out(f"    p{i + 1}: {ln[:200]}")
                        n += 1
                if n > 60:
                    break
