"""
Chile gas demand by sector + domestic production: source discovery.
Probes (from GitHub Actions; the editing sandbox can't reach these sites):
  - CNE Reporte Mensual (latest): every page mentioning natural gas, full text
  - INE "Electricidad, gas y agua" statistics (monthly gas distribution tables)
  - SEC statistics pages (gas de red sales by customer type)
  - CNE WordPress media library: Balance Nacional de Energia, gas statistics files
  - ENAP: annual reports / financial statements (Magallanes production)
"""
import io
import re

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)
GAS = re.compile(r"gas natural|\bGN\b|GNL|metanol|Methanex|Magallanes|ENAP|residencial|industrial|comercial|"
                 r"consumo|ventas|producci", re.I)


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
    found = []
    for h, t in re.findall(r'href="([^"]+)"[^>]*>(.{0,160}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I):
            if h.startswith("/"):
                h = base + h
            found.append((h, t))
    seen = set()
    res = []
    for h, t in found:
        if h not in seen:
            seen.add(h)
            res.append((h, t))
            out(f"   link: {t[:70]!r} -> {h[:200]}")
    return res


def pdf_gas_pages(content, full=True, max_pages=12):
    if content[:4] != b"%PDF":
        out("  not a PDF")
        return
    n = 0
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            if not re.search(r"gas natural", t, re.I):
                continue
            n += 1
            if n > max_pages:
                break
            out(f"  ---- page {i + 1}")
            for ln in t.splitlines()[:90]:
                if full or GAS.search(ln):
                    out("     ", ln[:220])


def excel_head(content, name):
    try:
        import pandas as pd
        x = pd.ExcelFile(io.BytesIO(content))
        out(f"  workbook {name}: sheets {x.sheet_names[:30]}")
        for s in x.sheet_names[:6]:
            df = pd.read_excel(x, sheet_name=s, header=None, nrows=40)
            out(f"   [{s}] shape {df.shape}")
            for _, row in df.head(40).iterrows():
                vals = [str(v)[:16] for v in row.tolist()[:14] if str(v) != "nan"]
                if vals:
                    out("      ", vals)
    except Exception as e:
        out("  excel read error", type(e).__name__, str(e)[:120])


# ------------------------------------------------------------------ CNE monthly report
out("=================== CNE Reporte Mensual (latest)")
for y, m in [(2026, 9), (2026, 8), (2026, 7)]:
    r = get(f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/RMensual_v{y}{m:02d}.pdf")
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        pdf_gas_pages(r.content)
        break

# ------------------------------------------------------------------ INE
out("\n=================== INE electricidad, gas y agua")
for u in ["https://www.ine.gob.cl/estadisticas/economia/energia",
          "https://www.ine.gob.cl/estadisticas/economia/energia/electricidad-gas-y-agua",
          "https://www.ine.gob.cl/estadisticas/economia/energia/electricidad-gas-y-agua/cuadros-estadisticos",
          "https://www.ine.gob.cl/estadisticas/economia/energia/gas"]:
    r = get(u)
    ls = links(r, r"gas|xlsx|xls|cuadro|boletin|energ", "https://www.ine.gob.cl")
    for h, t in ls:
        if re.search(r"\.xlsx?($|\?)", h, re.I) and re.search(r"gas", h + t, re.I):
            rr = get(h)
            if rr is not None and rr.status_code == 200:
                excel_head(rr.content, h)
            break

# ------------------------------------------------------------------ SEC
out("\n=================== SEC")
for u in ["https://www.sec.cl/", "https://www.sec.cl/estadisticas/", "https://www.sec.cl/gas-de-red/",
          "https://www.sec.cl/estadisticas-gas/", "https://www.sec.cl/combustibles/gas-de-red/",
          "https://www.sec.cl/informacion-estadistica/"]:
    links(get(u), r"estad|gas de red|gas-de-red|xlsx|xls|informe", "https://www.sec.cl")

# ------------------------------------------------------------------ CNE media library / pages
out("\n=================== CNE media search")
for q in ["Balance Nacional", "BNE", "gas natural", "consumo gas", "ventas gas", "estadisticas hidrocarburos",
          "produccion gas", "metanol"]:
    r = get("https://www.cne.cl/wp-json/wp/v2/media", params={"search": q, "per_page": 50})
    if r is not None and r.status_code == 200:
        for s in re.findall(r'"source_url":"([^"]+)"', r.text):
            s = s.replace("\\/", "/")
            if not re.search(r"\.(png|jpe?g|gif|webp)$", s, re.I):
                out(f"   [{q}] {s}")
for u in ["https://www.cne.cl/estadisticas/hidrocarburos/", "https://www.cne.cl/nuestros-servicios/reportes/",
          "https://www.cne.cl/nuestros-servicios/reportes/informacion-y-estadisticas/",
          "https://www.cne.cl/estadisticas/energia/", "https://www.cne.cl/estadisticas/"]:
    links(get(u), r"gas|balance|hidrocarb|xlsx|xls|estad", "https://www.cne.cl")

# ------------------------------------------------------------------ ENAP
out("\n=================== ENAP")
for u in ["https://www.enap.cl/", "https://www.enap.cl/pag/66/1290/memorias",
          "https://www.enap.cl/pag/63/1286/estados_financieros", "https://www.enap.cl/inversionistas"]:
    links(get(u), r"memoria|financ|razonado|inversion|produc|magallanes|pdf", "https://www.enap.cl")
