"""CEE probe 4: AGGM data endpoint discovery (candidate POSTs + Playwright 'Anwenden' capture). Writes probe_cee_out/aggm4.txt"""
import json, os, re
import requests
from playwright.sync_api import sync_playwright
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "probe_cee_out"); os.makedirs(OUT, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
L = []
B = "https://platform.aggm.at/vis-service/api/"
body = {"displayType": "line", "from": "2026-09-01T06:00:00", "granularity": "Day", "rangeType": "individual", "timeseries": ["ErmittelterEKVOesterreich"], "to": "2026-09-10T06:00:00"}
for p in ("ts/data", "ts/query", "ts/chart", "ts/timeseries", "ts/publication", "ts/values", "ts/search", "ts", "ts/table", "ts/export", "ts/tsdata", "ts/series", "ts/favorite", "ts/request"):
    for m in ("POST",):
        try:
            r = requests.request(m, B + p, json=body, headers={"User-Agent": UA}, timeout=30)
            L.append(f"{m} {p}: {r.status_code} {r.text[:200]!r}")
        except Exception as e: L.append(f"{p}: {e}")
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_context(user_agent=UA, viewport={"width": 1500, "height": 1200}).new_page()
    def on(r):
        if "vis-service" in r.url and r.request.method != "GET" or ("vis-service" in r.url and "attributes" not in r.url):
            L.append(f"[{r.status}] {r.request.method} {r.url[:300]}\n   post={(r.request.post_data or '')[:1500]}\n   resp={r.text()[:600] if r.status<400 else ''}")
    pg.on("response", on)
    pg.goto("https://platform.aggm.at/portal/visualisation/ts-publication?fav=Qo1", wait_until="domcontentloaded", timeout=60000)
    pg.wait_for_timeout(20000)
    for sel in ("text=Alle akzeptieren", "text=Anwenden"):
        try: pg.locator(sel).first.click(timeout=5000); pg.wait_for_timeout(8000); L.append("clicked " + sel)
        except Exception as e: L.append("click fail " + sel + type(e).__name__)
    # switch granularity to days via direct URL
    for u in ("https://platform.aggm.at/portal/visualisation/ts-publication?granularity=day&startDate=2026-09-01&endDate=2026-09-10&Consumption=EndConsumer&MarketArea=Austria",):
        pg.goto(u, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(15000)
        L.append("TEXT " + pg.inner_text("body")[:800].replace("\n", " | "))
        for sel in ("text=Anwenden",):
            try: pg.locator(sel).first.click(timeout=5000); pg.wait_for_timeout(8000); L.append("clicked " + sel)
            except Exception as e: L.append("click fail " + sel + type(e).__name__)
    b.close()
open(os.path.join(OUT, "aggm4.txt"), "w").write("\n".join(L))
