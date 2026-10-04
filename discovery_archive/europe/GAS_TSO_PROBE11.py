"""
Probe 11: Enagas statistical bulletin archive (all PDF links, text of Dec-2022 + latest bulletin demand tables);
Transgaz physical-flow table via form POST / Export; Gasgrid 'Gas consumption in Finland' xlsx; Plinacro site crawl. Prints only.
"""
import io
import re
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"
PAGER = EN + "/en/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico-gas/_jcr_content/responsiveGrid/container/filedownloadpaginati.nocache//enagas/components/content/"


def enagas():
    print("=== ENAGAS bulletin archive")
    pdfs = []
    for pg in range(1, 30):
        try:
            r = requests.get(PAGER, params={"page_": pg}, headers=H, timeout=60)
            ls = re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text)
            new = [x for x in ls if x not in pdfs]
            print(" page", pg, r.status_code, len(ls), "new", len(new))
            pdfs += new
            if not new:
                break
        except Exception as e:  # noqa: BLE001
            print(" page", pg, type(e).__name__)
    print(len(pdfs), "pdfs:", [p.split("/")[-1] for p in pdfs][:200])
    try:
        import pdfplumber
    except Exception:  # noqa: BLE001
        print("no pdfplumber")
        return
    pick = [p for p in pdfs if re.search(r"dic[a-z_]*22|dec[a-z_]*22", p, re.I)][:1] + pdfs[:1]
    for p in pick:
        r = requests.get(EN + p, headers=H, timeout=120)
        print("PDF", p.split("/")[-1], r.status_code, len(r.content))
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print(" pages", len(pdf.pages))
            n = 0
            for i, pg in enumerate(pdf.pages):
                t = pg.extract_text() or ""
                if re.search(r"(?i)\bdemand\b", t) and re.search(r"GWh", t):
                    print(f"--- page {i + 1}")
                    print(t[:1800])
                    n += 1
                    if n >= 4:
                        break


def transgaz():
    print("=== TRANSGAZ form")
    u = "https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197"
    s = requests.Session()
    s.headers.update({**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
    r0 = s.get(u, timeout=60)
    hidden = re.findall(r"<input[^>]+type=['\"]hidden['\"][^>]*>", r0.text)
    print("hidden inputs:", [h[:140] for h in hidden])
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"):
        from datetime import date
        a, b = date(2025, 1, 14), date(2025, 1, 15)
        data = {"data_start": a.strftime(fmt), "data_stop": b.strftime(fmt), "puncte": "SM", "um": "MW", "Submit": "Afiseaza", "grid_cmd": ""}
        for h in hidden:
            m = re.search(r"name=['\"]([^'\"]+)['\"][^>]*value=['\"]([^'\"]*)['\"]", h)
            if m and m.group(1) not in data:
                data[m.group(1)] = m.group(2)
        r = s.post(u, data=data, timeout=60)
        rows = re.findall(r"kgrRow|kgrAltRow", r.text)
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", "", r.text, flags=re.S)))
        i = txt.find("Data Code")
        print(fmt, r.status_code, len(r.text), "rows:", len(rows), "|", txt[i:i + 900] if i >= 0 else txt[:200])
    data.update({"Exportbtn": "Export", "IgnorePaging": "on"})
    data.pop("Submit", None)
    r = s.post(u, data=data, timeout=60)
    print("export", r.status_code, r.headers.get("content-type"), len(r.content), r.content[:200])


def gasgrid():
    print("=== GASGRID xlsx")
    u = "https://gasgrid.fi/wp-content/uploads/2026/09/Gas-consumption-in-Finland-7.9.2026-1.xlsx"
    r = requests.get(u, headers=H, timeout=60)
    print(r.status_code, len(r.content))
    if r.ok:
        import pandas as pd
        x = pd.ExcelFile(io.BytesIO(r.content))
        for sh in x.sheet_names:
            d = x.parse(sh, header=None)
            print(" sheet", sh, d.shape)
            print(d.head(12).to_string()[:2500])
            print(d.tail(4).to_string()[:1200])


def plinacro():
    print("=== PLINACRO crawl")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        for u in ("https://www.plinacro.hr/default.aspx?id=6", "https://www.plinacro.hr/default.aspx?id=109"):
            pg.goto(u, wait_until="domcontentloaded", timeout=45000)
            pg.wait_for_timeout(3000)
            seen = set()
            for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,80), e.href])"):
                if h not in seen and re.search(r"transport|tok|potro|podac|izvje|statist|operativ|transpar|alokac|nomin|bilanc|plin", t + h, re.I):
                    seen.add(h)
                    print("  ", t, "->", h[:130])
        b.close()


for fn in (enagas, transgaz, gasgrid, plinacro):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
