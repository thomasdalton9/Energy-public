"""
Probe 3: raw TSO daily national gas consumption for DK (Energinet Gasflow), PT (REN DataHub), LT (Amber Grid), FI (Gasgrid),
AT (AGGM), PL (Gaz-System), RO (Transgaz), SK, BG, HR. Direct API calls first, then headless-browser XHR capture for portals.
Prints only.
"""
import json
import re
import sys
from datetime import date, timedelta

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"


def show(label, fn):
    print("=" * 90, f"\n{label}", flush=True)
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print("  FAILED", type(e).__name__, str(e)[:300])


def dk():
    r = requests.get("https://api.energidataservice.dk/dataset/Gasflow", params={"limit": 5, "sort": "GasDay DESC"}, timeout=60)
    print(" status", r.status_code)
    print(" ", r.text[:900])
    r = requests.get("https://api.energidataservice.dk/meta/dataset/Gasflow", timeout=60)
    print(" meta", r.status_code, r.text[:1500])
    r = requests.get("https://api.energidataservice.dk/dataset/Gasflow", params={"start": "2021-01-01", "end": "2021-01-10", "limit": 20, "sort": "GasDay ASC"}, timeout=60)
    print(" 2021 sample", r.status_code, r.text[:500])


def pt():
    d = (date.today() - timedelta(days=5)).isoformat()
    for ep in ("GasConsumptionSupplyDaily", "GasConsumptionBreakdownDaily"):
        r = requests.get(f"https://servicebus.ren.pt/datahubapi/gas/{ep}", params={"culture": "en-US", "date": d}, timeout=60, headers={"User-Agent": UA})
        print(f" {ep} {d}:", r.status_code, r.text[:900])
    r = requests.get("https://servicebus.ren.pt/datahubapi/gas/GasConsumptionSupplyDaily", params={"culture": "en-US", "date": "2021-02-01"}, timeout=60, headers={"User-Agent": UA})
    print(" 2021-02-01:", r.status_code, r.text[:300])


def portals():
    from playwright.sync_api import sync_playwright
    pages = [("Amber Grid consumption", "https://ambergrid.lt/en/transmission-data"),
             ("Amber Grid root", "https://ambergrid.lt/en"),
             ("Gasgrid market info", "https://gasgrid.fi/en/gas-market/market-information/"),
             ("AGGM", "https://www.aggm.at/en/"),
             ("AGGM data", "https://www.aggm.at/en/data-and-figures/"),
             ("Gaz-System", "https://www.gaz-system.pl/en/"),
             ("Transgaz", "https://www.transgaz.ro/en"),
             ("Eustream", "https://www.eustream.sk/en"),
             ("Bulgartransgaz", "https://www.bulgartransgaz.bg/en/"),
             ("Plinacro", "https://www.plinacro.hr/default.aspx?id=1013")]
    with sync_playwright() as p:
        b = p.chromium.launch()
        for label, url in pages:
            ctx = b.new_context(user_agent=UA)
            pg = ctx.new_page()
            seen = []
            pg.on("response", lambda r: seen.append(f"[{r.status}] {r.request.method} {r.url[:170]}") if r.request.resource_type in ("xhr", "fetch") and not re.search(r"google|doubleclick|facebook|analytics|cookie|clarity|hotjar|demdex|dynatrace", r.url) else None)
            print("-" * 90, f"\n{label}: {url}", flush=True)
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=40000)
                pg.wait_for_timeout(5000)
            except Exception as e:  # noqa: BLE001
                print("  load problem", type(e).__name__)
            print("  title:", pg.title()[:90], "| url:", pg.url[:120])
            for s in seen[:12]:
                print("  xhr", s)
            n = 0
            for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,60), e.href])"):
                if re.search(r"consum|data|transparen|statistic|flow|gas-?market|market-?info|\.xlsx?|\.csv|zuzy|consum|verbrauch|download|export", h + " " + t, re.I):
                    print(f"  link {t!r} -> {h[:160]}")
                    n += 1
                    if n >= 14:
                        break
            ctx.close()
        b.close()


show("Denmark: Energinet Gasflow", dk)
show("Portugal: REN DataHub", pt)
show("Portals", portals)
sys.exit(0)
