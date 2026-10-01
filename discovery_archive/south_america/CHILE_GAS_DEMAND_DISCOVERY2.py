"""
Chile gas demand by sector, round 2 (round 1: CHILE_GAS_DEMAND_DISCOVERY.py found
no sector table in the CNE monthly report; the INE/SEC/ENAP pages guessed there 404'd
or only linked the menus). Probes:
  - CNE Estadisticas > Hidrocarburo page (xlsx statistics files, like the electricity page)
  - CNE media library: "Anuario" (Anuario Estadistico de Energia, BNE tables)
  - INE industry/energy pages (annual electricity-gas-water survey, IPEGA)
  - SEC home page links (gas statistics)
  - ENAP: financial statements + annual reports page, Magallanes page
  - Methanex investor pages (Chile methanol production by quarter)
Runs in GitHub Actions; the sites are blocked from the editing sandbox.
"""
import io
import re

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)


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


MENU = re.compile(r"/(estudios|tarificacion|normativas|participar)/|regiones\.ine|google\.|facebook|twitter|instagram|"
                  r"youtube|linkedin", re.I)


def links(r, pat, base="", show=80):
    if r is None or r.status_code != 200:
        return []
    res, seen = [], set()
    for h, t in re.findall(r'href="([^"]+)"[^>]*>(.{0,200}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if h.startswith("/"):
            h = base + h
        if h in seen or MENU.search(h) or not re.search(pat, h + " " + t, re.I):
            continue
        seen.add(h)
        res.append((h, t))
    for h, t in res[:show]:
        out(f"   link: {t[:70]!r} -> {h[:220]}")
    if len(res) > show:
        out(f"   ... {len(res) - show} more")
    return res


def excel_head(content, name, rows=30, sheets=8):
    try:
        import pandas as pd
        x = pd.ExcelFile(io.BytesIO(content))
        out(f"  workbook {name}: sheets {x.sheet_names[:40]}")
        for s in x.sheet_names[:sheets]:
            df = pd.read_excel(x, sheet_name=s, header=None, nrows=rows)
            out(f"   [{s}] shape {df.shape}")
            for _, row in df.iterrows():
                vals = [str(v)[:18] for v in row.tolist()[:16] if str(v) != "nan"]
                if vals:
                    out("      ", vals)
    except Exception as e:
        out("  excel read error", type(e).__name__, str(e)[:120])


def pdf_grep(content, pat, max_hits=80, pages=None):
    if content[:4] != b"%PDF":
        out("  not a PDF")
        return
    hits = 0
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            if pages and i not in pages:
                continue
            t = page.extract_text() or ""
            for ln in t.splitlines():
                if re.search(pat, ln, re.I):
                    out(f"    p{i + 1}: {ln[:220]}")
                    hits += 1
                    if hits >= max_hits:
                        return


# ------------------------------------------------------------------ CNE statistics: hidrocarburo
out("=================== CNE Estadisticas > Hidrocarburo")
r = get("https://www.cne.cl/estadisticas/hidrocarburo/")
ls = links(r, r"\.xlsx?|\.csv|\.pdf|gas|estad", "https://www.cne.cl", show=120)
for h, t in ls:
    if re.search(r"\.xlsx?($|\?)", h, re.I) and re.search(r"gas|GN|consumo|venta|producc|balance", h + t, re.I):
        rr = get(h)
        if rr is not None and rr.status_code == 200:
            excel_head(rr.content, h)

# ------------------------------------------------------------------ CNE media: Anuario / balance
out("\n=================== CNE media: anuario / balance")
for q in ["anuario", "Anuario Estadistico", "balance energia", "consumo sectorial", "gas de red", "metanol"]:
    r = get("https://www.cne.cl/wp-json/wp/v2/media", params={"search": q, "per_page": 60})
    if r is not None and r.status_code == 200:
        for s in re.findall(r'"source_url":"([^"]+)"', r.text):
            s = s.replace("\\/", "/")
            if not re.search(r"\.(png|jpe?g|gif|webp)$", s, re.I):
                out(f"   [{q}] {s}")

# ------------------------------------------------------------------ INE
out("\n=================== INE")
for u in ["https://www.ine.gob.cl/estadisticas-por-tema/industria-energia-y-construccion",
          "https://www.ine.gob.cl/estadisticas-por-tema/industria-energia-y-construccion/encuesta-anual-de-electricidad-gas-y-agua",
          "https://www.ine.gob.cl/estadisticas-por-tema/industria-energia-y-construccion/indice-de-produccion-industrial"]:
    links(get(u), r"gas|xlsx|xls|cuadro|boletin|electricidad|IPEGA|distribuci", "https://www.ine.gob.cl", show=50)

# ------------------------------------------------------------------ SEC
out("\n=================== SEC")
r = get("https://www.sec.cl/")
links(r, r"estad|gas|informe|dato", "https://www.sec.cl", show=80)

# ------------------------------------------------------------------ ENAP
out("\n=================== ENAP")
r = get("https://www.enap.cl/enap-transparente/estados-financieros-y-memorias")
ls = links(r, r"memoria|financ|razonado|\.pdf|20(2[1-6])", "https://www.enap.cl", show=80)
mem = [h for h, t in ls if re.search(r"memoria", h + t, re.I) and re.search(r"\.pdf|download|archivo", h, re.I)]
if mem:
    rr = get(mem[0])
    if rr is not None and rr.status_code == 200:
        pdf_grep(rr.content, r"gas natural.*(m3|m³|millones|MM)|producci[oó]n de gas|Magallanes.*gas|metanol|Methanex")
r = get("https://www.enap.cl/nuestras-operaciones/enap-en-magallanes")
if r is not None and r.status_code == 200:
    txt = re.sub(r"<[^>]+>", " ", r.text)
    for m in re.finditer(r"[^.]{0,200}(gas|m3|m³|metanol)[^.]{0,200}", txt, re.I):
        out("   txt:", re.sub(r"\s+", " ", m.group(0))[:300])
        break
links(get("https://www.enap.cl/informacion-financiera-2"), r"\.pdf|memoria|razonado|financ|gas", "https://www.enap.cl", 40)

# ------------------------------------------------------------------ Methanex
out("\n=================== Methanex")
for u in ["https://www.methanex.com/investor-relations/", "https://www.methanex.com/investor-relations/financial-reports/",
          "https://www.methanex.com/investor-relations/quarterly-results/"]:
    links(get(u), r"quarter|Q[1-4]|annual|report|\.pdf|MD&amp;A|mda", "https://www.methanex.com", 50)
