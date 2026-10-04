"""
Probe 13: Enagas bulletin pager (all /dam/ hrefs per pager page) + Dec-2022 bulletin demand table; Transgaz grid date handling
(export xlsx content, variants); Plinacro id=817 body and SUKAP portal XHR. Prints only (short).
"""
import io
import re
import sys

import pandas as pd
import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"


def ct(h):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h)).strip()


def enagas():
    print("=== ENAGAS pager 2")
    from playwright.sync_api import sync_playwright
    allp = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        pg.goto(EN + "/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/", wait_until="domcontentloaded", timeout=60000)
        pg.wait_for_timeout(4000)

        def grab(n):
            hs = pg.eval_on_selector_all("a[href*='/dam/']", "els => els.map(e => e.getAttribute('href'))")
            new = [h for h in hs if h not in allp]
            allp.extend(new)
            print(f" page {n}: {len(hs)} dam links, new {len(new)}", [h.split('/')[-1][:50] for h in hs][:5])
        grab(1)
        for n in range(2, 26):
            try:
                pg.locator(f".pagination a:text-is('{n}')").first.click(timeout=4000)
                pg.wait_for_timeout(2500)
                grab(n)
            except Exception as e:  # noqa: BLE001
                print(" page", n, "fail", type(e).__name__)
                try:
                    pg.locator("text=Next").first.click(timeout=3000)
                    pg.wait_for_timeout(2500)
                    grab(n)
                except Exception:  # noqa: BLE001
                    break
        print("pager html:", ct(pg.inner_html(".pagination") if pg.locator(".pagination").count() else "none")[:300])
        b.close()
    print(len(allp), "links")
    pick = [h for h in allp if re.search(r"dic.?22|dec.?22", h, re.I)]
    print("dec22:", pick)
    try:
        import pdfplumber
        for h in pick[:1]:
            r = requests.get(EN + h, headers=H, timeout=120)
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                for i, pgx in enumerate(pdf.pages[:6]):
                    t = pgx.extract_text() or ""
                    if "Evolution of gas demand" in t and "TOTAL" in t:
                        print(f"--- page {i + 1}\n{t[:1200]}")
                        break
    except Exception as e:  # noqa: BLE001
        print("pdf fail", type(e).__name__, str(e)[:200])


def transgaz():
    print("=== TRANSGAZ variants")
    u = "https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197"

    def first_dates(content, label):
        try:
            x = pd.read_excel(io.BytesIO(content), header=None)
            print(label, "xlsx", x.shape, x.iloc[:3].to_string()[:300].replace("\n", " | "), "| dates:", sorted(set(map(str, x.iloc[:, 0].dropna())))[:6])
        except Exception as e:  # noqa: BLE001
            print(label, "not xlsx", type(e).__name__, content[:100])
    for label, a, bb, ignore in (("1day", "2025-01-15", "2025-01-15", True), ("range", "2025-01-15", "2025-01-17", True), ("range-nopaging", "2025-01-15", "2025-01-17", False)):
        s = requests.Session()
        s.headers.update({**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
        r0 = s.get(u, timeout=60)
        vs = re.search(r"id='grid_viewstate'[^>]*value='([^']*)'", r0.text) or re.search(r"name='grid_viewstate'[^>]*value='([^']*)'", r0.text)
        data = {"data_start": a, "data_stop": bb, "puncte": "SM", "um": "MW", "Exportbtn": "Export", "grid_cmd": "", "grid_viewstate": vs.group(1) if vs else ""}
        if ignore:
            data["IgnorePaging"] = "on"
        r = s.post(u, data=data, timeout=60)
        first_dates(r.content, label + " export")
    # display variants
    s = requests.Session()
    s.headers.update({**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
    r0 = s.get(u, timeout=60)
    for label, extra in (("no viewstate", {}), ("grid_cmd=refresh", {"grid_cmd": "refresh"})):
        data = {"data_start": "2025-01-15", "data_stop": "2025-01-16", "puncte": "SM", "um": "MW", "Submit": "Afiseaza"}
        data.update(extra)
        r = s.post(u, data=data, timeout=60)
        body = re.sub(r"<script.*?</script>", "", r.text, flags=re.S)
        rows = re.findall(r"<tr[^>]*class=['\"]kgr(?:Row|AltRow)[^>]*>(.*?)</tr>", body, re.S)
        print(label, r.status_code, len(rows), [ct(x)[:40] for x in rows[:2]])


def plinacro():
    print("=== PLINACRO 817 / SUKAP")
    r = requests.get("https://www.plinacro.hr/default.aspx?id=817", headers=H, timeout=60)
    body = re.sub(r"<script.*?</script>|<style.*?</style>", "", r.text, flags=re.S)
    t = ct(body)
    i = t.rfind("Potrošnja distribucijskih sustava")
    print(t[i: i + 1800])
    for m in re.finditer(r'href="([^"]+)"[^>]*>', r.text):
        if re.search(r"xls|csv|pdf|UserDocs|download", m.group(1), re.I):
            print("   link", m.group(1)[:180])
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)

        def on(rs):
            if rs.request.resource_type in ("xhr", "fetch"):
                try:
                    body = rs.text()[:400].replace("\n", " ")
                except Exception:  # noqa: BLE001
                    body = ""
                print(f"   [{rs.status}] {rs.request.method} {rs.url[:160]} post={(rs.request.post_data or '')[:150]} resp={body}")
        pg.on("response", on)
        pg.goto("https://www.sukap.plinacro.hr/pub/", wait_until="networkidle", timeout=60000)
        pg.wait_for_timeout(5000)
        print("  sukap text:", pg.inner_text("body")[:700].replace("\n", " | "))
        for t_, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().slice(0,60), e.href])")[:40]:
            print("   link", t_, "->", h[:140])
        b.close()


for fn in (transgaz, plinacro, enagas):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
