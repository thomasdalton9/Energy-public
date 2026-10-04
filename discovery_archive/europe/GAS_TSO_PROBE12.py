"""
Probe 12: Enagas bulletin pager (playwright, click through pages, collect all PDF links) + text of the latest bulletin;
Transgaz grid table (ISO dates, rows, paging, export); Plinacro 'Potrosnja distribucijskih sustava' (id=817), id=816, sukap portal.
Prints only (kept short).
"""
import io
import re
import sys
from datetime import date

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"


def cell_text(h):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h)).strip()


def enagas():
    print("=== ENAGAS pager")
    from playwright.sync_api import sync_playwright
    pdfs = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        reqs = []
        pg.on("request", lambda r: reqs.append((r.method, r.url)) if "paginati" in r.url or "boletin" in r.url.lower() else None)
        pg.goto(EN + "/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/", wait_until="domcontentloaded", timeout=60000)
        pg.wait_for_timeout(4000)
        try:
            pg.locator("text=Accept all").first.click(timeout=3000)
        except Exception:  # noqa: BLE001
            pass

        def grab():
            for h in pg.eval_on_selector_all("a[href$='.pdf']", "els => els.map(e => e.getAttribute('href'))"):
                if h not in pdfs:
                    pdfs.append(h)
        grab()
        print("page1 pdfs", len(pdfs))
        for n in range(2, 30):
            try:
                loc = pg.locator(f".pagination a:text-is('{n}'), a[data-page='{n}'], li a:text-is('{n}')").first
                loc.click(timeout=4000)
                pg.wait_for_timeout(1500)
                before = len(pdfs)
                grab()
                print(" page", n, "pdfs", len(pdfs), "(+%d)" % (len(pdfs) - before))
            except Exception as e:  # noqa: BLE001
                print(" page", n, "click fail", type(e).__name__)
                break
        print("requests:", reqs[:12])
        b.close()
    print(len(pdfs), "pdfs total")
    print([x.split("/")[-1] for x in pdfs])
    try:
        import pdfplumber
        r = requests.get(EN + pdfs[0] if pdfs else EN + "/content/dam/enagas/en/files/gestion-tecnica-del-sistema/energy-data/publicaciones/boletin-estadistico-del-gas/Bolet%C3%ADn%20Estad%C3%ADstico_ago26_ingles.pdf", headers=H, timeout=120)
        print("PDF", r.status_code, len(r.content))
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print(" pages", len(pdf.pages))
            n = 0
            for i, pgx in enumerate(pdf.pages):
                t = pgx.extract_text() or ""
                if re.search(r"(?i)\bdemand\b", t) and re.search(r"GWh", t):
                    print(f"--- page {i + 1}")
                    print(t[:1500])
                    n += 1
                    if n >= 3:
                        break
    except Exception as e:  # noqa: BLE001
        print("pdf fail", type(e).__name__, str(e)[:200])


def transgaz():
    print("=== TRANSGAZ grid")
    u = "https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197"
    s = requests.Session()
    s.headers.update({**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
    r0 = s.get(u, timeout=60)
    vs = re.search(r"name='grid_viewstate'[^>]*value='([^']*)'", r0.text) or re.search(r"id='grid_viewstate'[^>]*value='([^']*)'", r0.text)
    for puncte in ("SM", "PM"):
        data = {"data_start": "2025-01-15", "data_stop": "2025-01-15", "puncte": puncte, "um": "MW", "Submit": "Afiseaza", "grid_cmd": "",
                "grid_viewstate": vs.group(1) if vs else ""}
        r = s.post(u, data=data, timeout=60)
        body = re.sub(r"<script.*?</script>", "", r.text, flags=re.S)
        rows = re.findall(r"<tr[^>]*class=['\"]kgr(?:Row|AltRow)[^>]*>(.*?)</tr>", body, re.S)
        print(puncte, r.status_code, len(r.text), "rows", len(rows))
        for row in rows[:40]:
            print("   ", cell_text(row)[:120])
        pager = re.findall(r"(?:Page|page|of|din)[^<]{0,40}", cell_text(body[-6000:]))
        print("   footer:", cell_text(body[body.find("kgrBottom"):][:1500])[:400] if "kgrBottom" in body else "none")
    data.update({"Exportbtn": "Export", "IgnorePaging": "on"})
    data.pop("Submit", None)
    r = s.post(u, data=data, timeout=60)
    print("export", r.status_code, r.headers.get("content-type"), len(r.content), r.content[:300])


def plinacro():
    print("=== PLINACRO")
    for u in ("https://www.plinacro.hr/default.aspx?id=817", "https://www.plinacro.hr/default.aspx?id=816", "https://www.sukap.plinacro.hr/pub/"):
        try:
            r = requests.get(u, headers=H, timeout=60)
            body = re.sub(r"<script.*?</script>|<style.*?</style>", "", r.text, flags=re.S)
            print(u, r.status_code, len(r.text))
            i = body.find("Potrošnja")
            print("  TEXT:", cell_text(body[max(0, i - 200): i + 1800])[:1500])
            for m in re.finditer(r'href="([^"]+\.(?:xlsx?|csv|pdf|zip)[^"]*)"[^>]*>([^<]{0,80})', r.text, re.I):
                print("   file", m.group(1)[:160], "|", m.group(2).strip())
            for m in re.finditer(r'<iframe[^>]+src="([^"]+)"', r.text):
                print("   iframe", m.group(1)[:200])
        except Exception as e:  # noqa: BLE001
            print(u, type(e).__name__)


for fn in (transgaz, plinacro, enagas):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
