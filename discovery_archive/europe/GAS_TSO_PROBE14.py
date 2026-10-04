"""
Probe 14: Enagas bulletin URL guessing (HEAD) for 2021-2026 + pager JS; Plinacro SUKAP consumption/flow XHR (playwright);
Transgaz export depth (2021). Prints only (short).
"""
import io
import re
import sys
from datetime import date

import pandas as pd
import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"
BASE = EN + "/content/dam/enagas/en/files/gestion-tecnica-del-sistema/energy-data/publicaciones/boletin-estadistico-del-gas/"
BASE_ES = EN + "/content/dam/enagas/es/ficheros/gestion-tecnica-del-sistema/energy-data/publicaciones/boletin-estadistico-del-gas/"


def enagas():
    print("=== ENAGAS guess")
    try:
        js = requests.get(EN + "/etc.clientlibs/enagas/clientlibs/components/filedownloadpagination.lc-d4ebae0b2f6418ce35d5fd115a33f336-lc.min.js", headers=H, timeout=60).text
        print("JS:", re.sub(r"\s+", " ", js)[:1800])
    except Exception as e:  # noqa: BLE001
        print("js fail", type(e).__name__)
    es = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
    en = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    found = []
    for y in (21, 22, 23, 24, 25):
        for m in range(12):
            hit = None
            for mon in dict.fromkeys([es[m], en[m]]):
                for suf in ("_ingles.pdf", "_ingles_rev1.pdf", "_espanol.pdf", ".pdf", "_ingles_rev.pdf"):
                    u = BASE + f"Bolet%C3%ADn%20Estad%C3%ADstico_{mon}{y}{suf}"
                    try:
                        r = requests.head(u, headers=H, timeout=30, allow_redirects=True)
                        if r.status_code == 200:
                            hit = u
                            break
                    except Exception:  # noqa: BLE001
                        pass
                if hit:
                    break
            if hit:
                found.append((y, m + 1, hit.split("/")[-1]))
    print("found", len(found), found)
    if found:
        try:
            import pdfplumber
            y, m, nm = [f for f in found if f[0] == 22 and f[1] == 12][0] if any(f[0] == 22 and f[1] == 12 for f in found) else found[0]
            r = requests.get(BASE + nm, headers=H, timeout=120)
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                for i, pg in enumerate(pdf.pages[:6]):
                    t = pg.extract_text() or ""
                    if "Evolution of gas demand" in t and "TOTAL" in t:
                        print(f"--- {nm} page {i + 1}\n{t[:1300]}")
                        break
        except Exception as e:  # noqa: BLE001
            print("pdf fail", type(e).__name__, str(e)[:200])


def transgaz():
    print("=== TRANSGAZ depth")
    u = "https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197"
    for a, b in (("2021-01-14", "2021-01-15"), ("2022-06-01", "2022-06-02"), ("2019-01-14", "2019-01-15")):
        s = requests.Session()
        s.headers.update({**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
        r0 = s.get(u, timeout=60)
        vs = re.search(r"id='grid_viewstate'[^>]*value='([^']*)'", r0.text)
        data = {"data_start": a, "data_stop": b, "puncte": "SM", "um": "MW", "Exportbtn": "Export", "IgnorePaging": "on", "grid_cmd": "", "grid_viewstate": vs.group(1) if vs else ""}
        r = s.post(u, data=data, timeout=60)
        try:
            x = pd.read_excel(io.BytesIO(r.content), header=None)
            print(a, x.shape, x.iloc[3:9].to_string().replace("\n", " | ")[:700])
        except Exception as e:  # noqa: BLE001
            print(a, "fail", type(e).__name__, r.content[:80])


def plinacro():
    print("=== PLINACRO SUKAP")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        for path in ("consumption", "flow"):
            pg = b.new_page(user_agent=UA)

            def on(rs, path=path):
                if rs.request.resource_type in ("xhr", "fetch") and "databaseParameter" not in rs.url and "currentMessages" not in rs.url:
                    try:
                        body = rs.text()[:700].replace("\n", " ")
                    except Exception:  # noqa: BLE001
                        body = ""
                    print(f"   [{rs.status}] {rs.request.method} {rs.url[:170]}\n       post={(rs.request.post_data or '')[:300]}\n       resp={body}")
            pg.on("response", on)
            print(" page", path)
            pg.goto("https://www.sukap.plinacro.hr/pub/" + path, wait_until="networkidle", timeout=60000)
            pg.wait_for_timeout(6000)
            print("  text:", pg.inner_text("body")[:1200].replace("\n", " | ")[-700:])
            pg.close()
        b.close()


for fn in (enagas, transgaz, plinacro):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
