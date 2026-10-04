"""
Probe 8: follow-ups from probe 7.
 - Enagas gas statistical bulletin file list (AJAX fragment) + other energy-data pages
 - Transgaz new-tabel-transparenta-masuratori_en.php?poz=N (the physical-flows table fragment)
 - Gaz-System MIR menu item 'Faktyczna ilosc przeslanego gazu' (XHR capture)
 - Bulgartransgaz historical data page links
 - Gasgrid: iframes / XHR
 - Plinacro: all links on id=109
Prints only.
"""
import re
import sys

import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}


def show(url, n=1500, **kw):
    try:
        r = requests.get(url, timeout=40, headers=kw.pop("headers", H), **kw)
        print(f"GET {url[:150]} -> {r.status_code} len={len(r.text)}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"GET {url[:150]} -> {type(e).__name__}")


def hrefs(txt, pat=r"\.(xlsx?|csv|zip|pdf)|download|jcr_content|dam/"):
    seen = set()
    for m in re.finditer(r'href="([^"]+)"[^>]*>([^<]{0,80})', txt):
        if re.search(pat, m.group(1), re.I) and m.group(1) not in seen:
            seen.add(m.group(1))
            print("   href", m.group(1)[:200], "|", m.group(2).strip()[:60])


def static():
    print("=== ENAGAS bulletin fragment")
    for u in ("https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico-gas/_jcr_content/responsiveGrid/container/filedownloadpaginati.nocache.html/enagas/components/content/filedownloadpagination",):
        r = show(u)
        if r is not None:
            print(r.text[:300])
    # the exact fragment url is cut in the log; fetch the bulletin page and extract
    r = show("https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/")
    if r is not None:
        for m in re.finditer(r'(/content/enagas[^"\' ]+nocache[^"\' ]*)', r.text):
            print("   nocache", m.group(1)[:250])
        for m in re.finditer(r'data-[a-z-]*(?:url|path|resource)[a-z-]*="([^"]+)"', r.text):
            print("   data-attr", m.group(1)[:200])
    print("=== TRANSGAZ fragment")
    for poz in (197, 196, 198, 1, 100, 150, 200, 210):
        r = show(f"https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz={poz}",
                 headers={**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
        if r is not None and r.status_code == 200:
            body = re.sub(r"<script.*?</script>", "", r.text, flags=re.S)
            print("   ", re.sub(r"\s+", " ", body)[:1200])
            hrefs(body)
    print("=== BULGARTRANSGAZ historical")
    for u in ("https://bulgartransgaz.bg/en/pages/istoricheski-danni-45.html",):
        r = show(u)
        if r is not None:
            hrefs(r.text, r"\.(xlsx?|csv|zip|pdf)|upload|files")
    print("=== PLINACRO links")
    for u in ("https://www.plinacro.hr/default.aspx?id=109", "https://www.plinacro.hr/default.aspx?id=1186"):
        r = show(u)
        if r is not None:
            for m in re.finditer(r'href="([^"]+)"[^>]*>([^<]{0,90})', r.text):
                h = m.group(1)
                if re.search(r"default\.aspx\?id=|\.(xlsx?|pdf|csv)|UserDocs", h, re.I):
                    print("   ", h[:140], "|", m.group(2).strip()[:70])


def browser():
    with sync_playwright() as p:
        b = p.chromium.launch()
        # Gaz-System MIR
        ctx = b.new_context(user_agent=UA)
        pg = ctx.new_page()

        def on(r):
            if r.request.resource_type in ("xhr", "fetch") and "gaz-system" in r.url:
                try:
                    body = r.text()[:600].replace("\n", " ")
                except Exception:  # noqa: BLE001
                    body = ""
                print(f"   [{r.status}] {r.request.method} {r.url[:200]}\n       post={(r.request.post_data or '')[:500]}\n       resp={body}")
        pg.on("response", on)
        print("=== GAZ-SYSTEM MIR")
        pg.goto("https://swi.gaz-system.pl/mir/", wait_until="domcontentloaded", timeout=45000)
        pg.wait_for_timeout(5000)
        for t in ("Faktyczna ilość przesłanego gazu", "Faktyczna ilość"):
            loc = pg.get_by_text(t).first
            try:
                if loc.count():
                    print("  clicking", t)
                    loc.click(timeout=5000)
                    pg.wait_for_timeout(9000)
                    break
            except Exception as e:  # noqa: BLE001
                print("  click fail", type(e).__name__)
        print("  url:", pg.url)
        print("  text:", pg.inner_text("body")[:900].replace("\n", " | "))
        for t_, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().slice(0,60), e.href])")[:60]:
            print("   link", t_, "->", h[:160])
        ctx.close()
        # Gasgrid
        ctx = b.new_context(user_agent=UA)
        pg = ctx.new_page()

        def on2(r):
            if r.request.resource_type in ("xhr", "fetch", "document") and not re.search(r"google|cookiebot|hcaptcha|facebook|youtube|admin-ajax|wp-content|wp-json/wp/", r.url):
                print(f"   [{r.status}] {r.request.method} {r.url[:220]}")
        pg.on("response", on2)
        print("=== GASGRID")
        pg.goto("https://gasgrid.fi/en/gas-business/what-is-flowing-in-our-pipelines/", wait_until="domcontentloaded", timeout=45000)
        pg.wait_for_timeout(12000)
        for fr in pg.frames:
            print("  frame", fr.url[:200])
        print("  iframes:", pg.eval_on_selector_all("iframe", "els => els.map(e => e.src || e.dataset.src)"))
        txt = pg.inner_text("body")
        i = txt.find("flowing")
        print("  text:", txt[i:i + 1500].replace("\n", " | "))
        for t_, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().slice(0,60), e.href])"):
            if re.search(r"xlsx|csv|pdf|data|consum|power|bi\.|datahub|api", h + t_, re.I):
                print("   link", t_, "->", h[:170])
        ctx.close()
        b.close()


for fn in (static, browser):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
