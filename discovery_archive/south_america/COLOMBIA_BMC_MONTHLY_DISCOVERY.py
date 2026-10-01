"""
Colombia gas demand by sector from the Gestor del Mercado de Gas (Bolsa
Mercantil de Colombia) MONTHLY reports - public PDFs, unlike the BI Gas
dashboard which needs a login + reCAPTCHA.
  https://www.bmcbec.com.co/informes/informes-mensuales
1. List every report link on that page (raw HTML and rendered via
   Playwright, following pagination / "ver mas"), with titles.
2. Download the newest report and print, page by page, the lines that
   carry a demand-by-sector table (termico/termoelectrico, industrial,
   residencial/comercial, GNV/vehicular, refineria, petroquimica, total)
   plus the month header, so a parser can be written against it.
Not reachable from the editing sandbox.
"""
import io
import re
from urllib.parse import urljoin

import requests

BASE = "https://www.bmcbec.com.co"
PAGE = BASE + "/informes/informes-mensuales"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
SECTOR = re.compile(r"t[eé]rmic|termoel|industri|residen|comerci|gnv|vehicul|refiner|petroqu|compres|demanda total|total demanda", re.I)


def out(*a):
    print(*a, flush=True)


links = {}
r = requests.get(PAGE, headers=H, timeout=T)
out(f"GET {PAGE} -> {r.status_code} {len(r.text)} chars")
for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
    t = re.sub(r"<[^>]+>|\s+", " ", txt).strip()
    if re.search(r"\.pdf|informe|mensual|download|sites/default/files", href, re.I):
        links[urljoin(BASE, href)] = t

from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(user_agent=H["User-Agent"])
    pg.goto(PAGE, timeout=90000, wait_until="load")
    pg.wait_for_timeout(6000)
    out("rendered title:", pg.title())
    for _ in range(6):
        for a in pg.locator("a").all():
            try:
                href = a.get_attribute("href") or ""
                t = (a.inner_text() or "").strip().replace("\n", " ")
            except Exception:
                continue
            if re.search(r"\.pdf|informe-mensual|/informe|sites/default/files|download", href, re.I):
                links[urljoin(BASE, href)] = t or links.get(urljoin(BASE, href), "")
        nxt = pg.locator("a[rel=next], li.pager__item--next a, a:has-text('Siguiente'), a:has-text('Ver más')")
        if nxt.count() == 0:
            break
        try:
            nxt.first.click(timeout=5000)
            pg.wait_for_timeout(3000)
        except Exception:
            break
    b.close()

out(f"\n{len(links)} candidate links:")
for u, t in sorted(links.items()):
    out(f"  {t[:90]!r} -> {u}")

pdfs = [u for u in links if u.lower().split("?")[0].endswith(".pdf")]
if not pdfs:
    # report pages may hold the PDF one level down
    for u in [u for u in links if "informe" in u.lower()][:8]:
        rr = requests.get(u, headers=H, timeout=T)
        for h in re.findall(r'href="([^"]+\.pdf[^"]*)"', rr.text, re.I):
            pdfs.append(urljoin(BASE, h))
            out(f"  pdf under {u}: {urljoin(BASE, h)}")
out(f"\n{len(pdfs)} pdf links")
cands = sorted(set(pdfs), key=lambda u: re.findall(r"20\d\d", u)[-1:] + [u], reverse=True)

import pdfplumber
for u in cands[:2]:
    out(f"\n==================== {u}")
    c = requests.get(u, headers=H, timeout=T).content
    if c[:4] != b"%PDF":
        out("  not a PDF:", c[:80])
        continue
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = text.splitlines()
            hit = [ln for ln in lines if SECTOR.search(ln)]
            if len(hit) >= 3:
                out(f"  --- page {i + 1}: {lines[0][:120] if lines else ''}")
                for ln in lines:
                    if SECTOR.search(ln) or re.search(r"(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[a-z]*[-/ ]?\d{2,4}|GBTUD|Gbtud|MPCD", ln, re.I):
                        out("     ", ln[:200])
                tables = page.extract_tables()
                for tb in tables[:2]:
                    out("     TABLE rows:", len(tb))
                    for row in tb[:14]:
                        out("       ", [str(x)[:18] if x else "" for x in row][:14])
