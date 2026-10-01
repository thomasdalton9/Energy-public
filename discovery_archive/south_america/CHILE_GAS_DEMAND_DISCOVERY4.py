"""
Chile gas demand by sector, round 4. Round 3 found:
  - Consumo_mensual_GN (CNE, monthly consumption by sector) ends Oct 2018: no 2021+ data
  - Produccion_combustibles-junio-2024.xlsx: monthly ENAP + CEOP gas production,
    Jan 2018 - Jun 2024 (latest version in the CNE media library)
  - Anuario Estadistico de Energia 2021 (only edition from 2021 on) has a
    "Consumo final de energia" section (p152-160, annual, Tcal)
  - importaciones-web.xlsx 'GAS NATURAL GASEOSO' has an 'IMPORTACION USO ...' column
This round prints: the Anuario consumption pages; the import workbook's gas headers
and recent rows; the Ministry of Energy BNE pages; gas lines in ENAP's 2025
integrated report. Runs in GitHub Actions.
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


# ------------------------------------------------------------------ Anuario 2021 consumption pages
out("=================== Anuario 2021 p150-162")
r = get("https://www.cne.cl/wp-content/uploads/2022/07/AnuarioEstadisticoEnergia2021.pdf")
if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        for i in range(149, min(162, len(pdf.pages))):
            out(f"  ---- page {i + 1}")
            for ln in (pdf.pages[i].extract_text() or "").splitlines()[:80]:
                out("     ", ln[:200])
        for i in (98, 99, 100):  # LNG terminals section
            out(f"  ---- page {i + 1}")
            for ln in (pdf.pages[i].extract_text() or "").splitlines()[:60]:
                out("     ", ln[:200])

# ------------------------------------------------------------------ imports workbook: gas sheets
out("\n=================== importaciones-web gas sheets")
r = get("https://www.cne.cl/wp-content/uploads/2026/06/importaciones-web.xlsx")
if r is not None and r.status_code == 200:
    x = pd.ExcelFile(io.BytesIO(r.content))
    for s in ("GAS NATURAL GASEOSO", "GAS NATURAL LICUADO"):
        df = pd.read_excel(x, sheet_name=s, header=None)
        out(f"  [{s}] shape {df.shape}")
        for i in range(0, 12):
            out(f"     {i}: {[str(v)[:40] for v in df.iloc[i].tolist()]}")
        for i in range(len(df) - 30, len(df)):
            vals = [str(v)[:14] for v in df.iloc[i].tolist()]
            if any(v != "nan" for v in vals):
                out(f"     {i}: {vals}")

# ------------------------------------------------------------------ Ministry of Energy BNE
out("\n=================== Ministerio de Energia BNE")
for u in ["https://energia.gob.cl/", "https://www.energia.gob.cl/balance-nacional-de-energia",
          "https://energia.gob.cl/mini-sitio/balance-nacional-de-energia",
          "https://www.energia.gob.cl/documentos?search=balance"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        for h, t in re.findall(r'href="([^"]+)"[^>]*>(.{0,160}?)</a', r.text, re.S):
            t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
            if re.search(r"balance|BNE|\.xlsx?", h + " " + t, re.I):
                out(f"   link: {t[:80]!r} -> {h[:200]}")

# ------------------------------------------------------------------ ENAP 2025 integrated report
out("\n=================== ENAP Reporte Integrado 2025: gas lines")
r = get("https://www.enap.cl/files/get/2995")
if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
    n = 0
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        for i, p in enumerate(pdf.pages):
            for ln in (p.extract_text() or "").splitlines():
                if re.search(r"\bgas\b", ln, re.I) and re.search(r"m3|m³|MMm|millones|MMpc|Mm3|bcf|pies", ln, re.I):
                    out(f"    p{i + 1}: {ln[:200]}")
                    n += 1
            if n > 80:
                break
    out(f"  {n} lines")
