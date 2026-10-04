"""
Probe 4: raw daily gas consumption pages/APIs for LT (Amber Grid open data), FI (Gasgrid Datahub), AT (AGGM platform / data monitor),
PL (Gaz-System data transparency), BG (Bulgartransgaz operational + historical data), HR (Plinacro), RO (Transgaz physical flows /
technological consumption), SK. Loads each page headless, logs XHR/JSON and every data-like link, plus page text.
Prints only.
"""
import re
import sys

from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
PAGES = [
    ("LT Amber Grid open data", "https://ambergrid.lt/en/for-clients/open-data/650"),
    ("LT Amber Grid capacity and gas flow", "https://ambergrid.lt/en/for-clients/services/capacity-and-gas-flow/639"),
    ("FI Gasgrid Datahub", "https://katso.azurewebsites.net/"),
    ("FI Gasgrid market data", "https://gasgrid.fi/en/gas-business/finlands-gas-markets/"),
    ("AT AGGM data monitor", "https://www.aggm.at/en/transparency/data-monitor/"),
    ("AT AGGM platform", "https://platform.aggm.at/portal/visualisation/map"),
    ("PL Gaz-System data transparency", "https://www.gaz-system.pl/en/for-customers/market-information/data-transparency.html"),
    ("BG Bulgartransgaz operational", "https://bulgartransgaz.bg/en/pages/operational-data-185.html"),
    ("BG Bulgartransgaz historical", "https://bulgartransgaz.bg/en/pages/istoricheski-danni-45.html"),
    ("HR Plinacro transparency", "https://www.plinacro.hr/default.aspx?id=109"),
    ("RO Transgaz physical flows", "https://www.transgaz.ro/en/clients/operational-data/physical-flows"),
    ("RO Transgaz tech consumption", "https://www.transgaz.ro/en/clients/operational-data/technological-consumption"),
    ("SK Eustream business data", "https://www.eustream.sk/en/transparency/business-operational-data/"),
]
SKIP = re.compile(r"google|doubleclick|facebook|analytics|cookie|clarity|hotjar|demdex|dynatrace|visualstudio|office\.net|monitor\.azure|powerbi\.com/13|approvedResources", re.I)

with sync_playwright() as p:
    b = p.chromium.launch()
    for label, url in PAGES:
        ctx = b.new_context(user_agent=UA)
        pg = ctx.new_page()
        seen = []

        def on(r, seen=seen):
            if r.request.resource_type in ("xhr", "fetch") and not SKIP.search(r.url):
                try:
                    body = r.text()[:160].replace("\n", " ")
                except Exception:  # noqa: BLE001
                    body = ""
                seen.append(f"[{r.status}] {r.request.method} {r.url[:170]} :: {body}")
        pg.on("response", on)
        print("=" * 100, f"\n{label}: {url}", flush=True)
        try:
            pg.goto(url, wait_until="domcontentloaded", timeout=40000)
            pg.wait_for_timeout(7000)
        except Exception as e:  # noqa: BLE001
            print("  load problem", type(e).__name__)
        print("  title:", pg.title()[:90], "| url:", pg.url[:120])
        for s_ in seen[:14]:
            print("  xhr", s_)
        n = 0
        for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,70), e.href])"):
            if re.search(r"\.(xlsx?|csv|zip|json)(\?|$)|consum|verbrauch|zuzy|potrebl|dnevn|daily|physical|flow|download|export|history|istorich|archive", h + " " + t, re.I):
                print(f"  link {t!r} -> {h[:170]}")
                n += 1
                if n >= 18:
                    break
        try:
            print("  text:", pg.inner_text("body")[:500].replace("\n", " | "))
        except Exception:  # noqa: BLE001
            pass
        ctx.close()
    b.close()
sys.exit(0)
