"""
Round 3 of the gas/power source search (all public, no keys):

A. Trinidad 2025 - no full-year 2025 bulletin is linked from the bulletin
   category. Look for 2025 editions (monthly or year-to-date) via the
   WordPress posts/pages REST API (media API is 401), the sitemaps, and
   site search; open any 2025 workbook found and list its months.
B. Chile - CNE Reporte Mensual: dump every table on the generation pages
   (generation by technology) and the imports pages. CEN's own monthly
   "Reporte Energetico SEN" PDF: generation by technology / fuel tables.
C. Ecuador - Petroecuador monthly report Jan-Aug 2026 (id 3889), annual
   2025 (id 3860), monthly national dispatches (id 3292): every line
   mentioning gas, with its table header.
D. Bolivia - MHE quarterly bulletin via Chromium (python requests fails
   the site's certificate chain), INE hydrocarbons statistical tables.
Not reachable from the editing sandbox.
"""
import io
import re
from urllib.parse import urljoin

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
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:30]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:120]}")
        return None


def pdf_dump(content, page_pat, line_pat=None, max_lines=120, tables=True, max_pages=12):
    if content[:4] != b"%PDF":
        out("  not a PDF:", content[:80])
        return
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages; p1: {(pdf.pages[0].extract_text() or '')[:150]!r}")
        shown = 0
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not re.search(page_pat, text, re.I):
                continue
            shown += 1
            if shown > max_pages:
                break
            out(f"  --- page {i + 1}")
            for ln in text.splitlines()[:max_lines]:
                if line_pat is None or re.search(line_pat, ln, re.I):
                    out("     ", ln[:210])
            if tables:
                for tb in page.extract_tables()[:3]:
                    out(f"     TABLE {len(tb)} rows")
                    for row in tb[:30]:
                        out("       ", [str(x)[:16].replace("\n", " ") if x else "" for x in row][:16])


# ============================================================ A. Trinidad 2025
out("=================== A. TRINIDAD 2025")
SITE = "https://www.energy.gov.tt"
cands = set()
for kind in ("posts", "pages"):
    for term in ("bulletin", "consolidated", "2025"):
        r = get(f"{SITE}/wp-json/wp/v2/{kind}", params={"search": term, "per_page": 100})
        if r is None or r.status_code != 200:
            continue
        try:
            for p in r.json():
                title = re.sub(r"<[^>]+>", "", p.get("title", {}).get("rendered", ""))
                body = p.get("content", {}).get("rendered", "")
                files = re.findall(r'href="([^"]+\.(?:xlsx?|pdf))"', body, re.I)
                if re.search(r"2025", title + " ".join(files)):
                    out(f"  [{kind}/{term}] {title[:90]!r} {p.get('link')}")
                    for f in files:
                        out("       ", f)
                        cands.add(f)
        except Exception as e:
            out("  json err", e)
for sm in ("/wp-sitemap.xml", "/sitemap_index.xml", "/sitemap.xml"):
    r = get(SITE + sm)
    if r is None or r.status_code != 200:
        continue
    subs = re.findall(r"<loc>([^<]+)</loc>", r.text)
    out(f"  {sm}: {len(subs)} entries")
    for s in subs[:60]:
        if re.search(r"post|page|attachment|media", s):
            rr = get(s)
            if rr is None or rr.status_code != 200:
                continue
            for loc in re.findall(r"<loc>([^<]+)</loc>", rr.text):
                if re.search(r"bulletin", loc, re.I) and re.search(r"2025|2026", loc):
                    out("       sitemap hit:", loc)
                    if re.search(r"\.(xlsx?|pdf)$", loc, re.I):
                        cands.add(loc)
                    else:
                        r3 = get(loc)
                        if r3 is not None and r3.status_code == 200:
                            for f in re.findall(r'href="([^"]+\.(?:xlsx?|pdf))"', r3.text, re.I):
                                if re.search(r"bulletin", f, re.I):
                                    out("           ", f)
                                    cands.add(f)
    break
r = get(SITE + "/", params={"s": "bulletin 2025"})
if r is not None and r.status_code == 200:
    for h in sorted(set(re.findall(r'href="(https://www\.energy\.gov\.tt/[^"]+)"', r.text))):
        if re.search(r"bulletin|2025", h, re.I):
            out("   search:", h)
for f in sorted(cands):
    if re.search(r"\.xlsx?$", f, re.I) and "2025" in f:
        r = get(f)
        if r is None or r.status_code != 200:
            continue
        try:
            xl = pd.ExcelFile(io.BytesIO(r.content))
            sheet = next((s for s in xl.sheet_names if s.strip().replace(" ", "") in ("3A,3B", "3A3B")), None)
            df = xl.parse(sheet, header=None) if sheet else None
            months = sorted({pd.Timestamp(v).strftime("%Y-%m") for row in df.head(10).itertuples(index=False)
                             for v in row if hasattr(v, "year")}) if df is not None else None
            out(f"  {f.rsplit('/', 1)[-1]}: sheet {sheet!r} months {months}")
        except Exception as e:
            out(f"  {f}: ERR {e}")

# ============================================================ B. Chile
out("\n=================== B. CHILE: CNE Reporte Mensual Sep 2026 - generation & imports pages")
r = get("https://www.cne.cl/wp-content/uploads/2026/09/RMensual_v202609.pdf")
if r is not None and r.status_code == 200:
    pdf_dump(r.content, r"Generación Eléctrica SEN|generación por tecnología|Importaciones|GWh", max_lines=80)
out("\n=================== B. CHILE: CEN Reporte Energetico SEN")
mon = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
got = None
for y, m in [(2026, 9), (2026, 8), (2026, 7), (2026, 6)]:
    for up in [(y, m), (y, m + 1 if m < 12 else 1)]:
        u = f"https://www.coordinador.cl/wp-content/uploads/{up[0]}/{up[1]:02d}/CEN_Reporte_Energetico_SEN_{mon[m - 1]}{str(y)[2:]}.pdf"
        rr = get(u)
        if rr is not None and rr.status_code == 200 and rr.content[:4] == b"%PDF":
            got = rr
            break
    if got:
        break
if got:
    pdf_dump(got.content, r"Gas Natural|GNL|combustible|tecnolog", max_lines=90, max_pages=10)
for y, m in [(2021, 1), (2022, 1), (2024, 1)]:
    get(f"https://www.coordinador.cl/wp-content/uploads/{y}/{m + 1:02d}/CEN_Reporte_Energetico_SEN_{mon[m - 1]}{str(y)[2:]}.pdf")
r = get("https://www.coordinador.cl/reportes-y-estadisticas/")
if r is not None and r.status_code == 200:
    for h in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
        if re.search(r"reporte|estad|generac|combust|\.xlsx|\.csv", h, re.I):
            out("   link:", h[:170])

# ============================================================ C. Ecuador
out("\n=================== C. ECUADOR: Petroecuador gas tables")
for i, label in ((3889, "monthly Jan-Aug 2026"), (3860, "annual 2025"), (3292, "despachos mensuales")):
    out(f"\n  ##### id={i} ({label})")
    r = get(f"https://www.eppetroecuador.ec/wp-content/plugins/download-monitor/download.php?id={i}")
    if r is None or r.status_code != 200:
        continue
    ct = r.headers.get("content-type", "")
    if r.content[:4] == b"%PDF":
        pdf_dump(r.content, r"GAS NATURAL|AMISTAD|GNL",
                 r"DESPACHO|Cifras|Enero|ENERO|GAS|AMISTAD|Sector|SECTOR|Eléctric|Industri|Producto|PRODUCTO",
                 max_lines=200, tables=False, max_pages=10)
    else:
        try:
            for name, df in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
                out(f"   sheet {name!r} {df.shape}")
                for row in df.head(15).itertuples(index=False):
                    out("      ", [str(x)[:14] for x in row if str(x) != "nan"][:14])
        except Exception as e:
            out(f"   not pdf/xlsx ({ct}): {e}")

# ============================================================ D. Bolivia
out("\n=================== D. BOLIVIA")
r = get("https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/hidrocarburo-cuadros-estadisticos/")
xl_links = []
if r is not None and r.status_code == 200:
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(r"\.xlsx?|\.csv|gas", h + t, re.I):
            out(f"   {t[:100]!r} -> {h}")
            if re.search(r"\.xlsx?$", h, re.I):
                xl_links.append((h, t))
for h, t in [x for x in xl_links if re.search(r"gas|consum|venta|comerc|mercado", x[0] + x[1], re.I)][:4]:
    rr = get(urljoin("https://www.ine.gob.bo/", h))
    if rr is None or rr.status_code != 200:
        continue
    try:
        for name, df in list(pd.read_excel(io.BytesIO(rr.content), sheet_name=None, header=None).items())[:3]:
            out(f"   ## {t[:60]!r} sheet {name!r} {df.shape}")
            for row in df.head(25).itertuples(index=False):
                out("      ", [str(x)[:14] for x in row if str(x) != "nan"][:14])
    except Exception as e:
        out("   xls err", e)

out("\n  -- MHE quarterly bulletin via Chromium")
try:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=H["User-Agent"])
        resp = ctx.request.get("https://www.mhe.gob.bo/wp-content/uploads/2025/12/Boletin-trimestral-3T_2025-Intranet.pdf",
                               timeout=120000)
        out("   status", resp.status, len(resp.body()))
        if resp.ok:
            pdf_dump(resp.body(), r"gas natural|mercado interno|termoel|GNV",
                     None, max_lines=70, max_pages=8)
        pg = ctx.new_page()
        pg.goto("https://www.mhe.gob.bo/", timeout=90000)
        for a in pg.locator("a").all()[:400]:
            try:
                h = a.get_attribute("href") or ""
                t = (a.inner_text() or "").strip()
            except Exception:
                continue
            if re.search(r"bolet|estad|\.pdf", h + t, re.I):
                out(f"   {t[:70]!r} -> {h}")
        b.close()
except Exception as e:
    out("   playwright err:", type(e).__name__, str(e)[:200])
