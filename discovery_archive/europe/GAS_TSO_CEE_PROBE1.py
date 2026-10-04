"""CEE probe 1: AGGM ts-publication data requests, NET4GAS CAMS extranet, Amber Grid Power BI querydata. Prints only."""
import re, sys
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
SKIP = re.compile(r"google|doubleclick|facebook|analytics|cookie|clarity|hotjar|demdex|dynatrace|office\.net|\.svg|\.woff|fonts|\.png|\.jpg|\.css|\.js(\?|$)", re.I)
PAGES = [
    ("AGGM Endkunden Ost", "https://platform.aggm.at/portal/visualisation/ts-publication?src=map&granularity=hour&startDate=2026-09-01&Consumption=EndConsumer", 30000),
    ("AGGM fav", "https://platform.aggm.at/portal/visualisation/ts-publication?fav=Qo1", 30000),
    ("NET4GAS CAMS", "https://extranet.cams.net4gas.cz/", 12000),
    ("AmberGrid open data", "https://ambergrid.lt/en/for-clients/open-data/650", 40000),
]
with sync_playwright() as p:
    b = p.chromium.launch()
    for label, url, wait in PAGES:
        pg = b.new_context(user_agent=UA).new_page()
        def on(r):
            if SKIP.search(r.url) and "powerbi" not in r.url: return
            if r.request.resource_type not in ("xhr", "fetch", "document"): return
            try: body = r.text()[:400].replace("\n", " ")
            except Exception: body = ""
            print(f"  [{r.status}] {r.request.method} {r.url[:230]}\n     post={(r.request.post_data or '')[:1500]}\n     resp={body}", flush=True)
        pg.on("response", on)
        print("=" * 90, "\n", label, flush=True)
        try:
            pg.goto(url, wait_until="domcontentloaded", timeout=45000)
            pg.wait_for_timeout(wait)
        except Exception as e: print("  load problem", type(e).__name__)
        print("  title:", pg.title()[:80], pg.url[:150])
        for fr in pg.frames: print("  frame", fr.url[:250])
        try: print("  text:", pg.inner_text("body")[:600].replace("\n", " | "))
        except Exception: pass
        for t, h in pg.eval_on_selector_all("a[href]", "els=>els.map(e=>[e.innerText.trim().slice(0,60),e.href])")[:60]:
            if re.search(r"xlsx?|csv|download|export|consum|data", h+t, re.I): print("  link", t, h[:160])
    b.close()
