"""
Probe 7 (Spain/PL/RO/HR/FI/BG operator pages): (1) Gassco FLOW (umm.gassco.no) past the disclaimer: XHR/JSON, history links, and the gassco.eu realTimeGraphData feed;
(2) AGGM ts API: full attributes + the data request the page makes when a series is selected;
(3) NET4GAS extranet.cams.net4gas.cz (current transmission system data);
(4) Enagas publications / physical-parameter pages for Excel history (Spain 2021-22);
(5) Transgaz statistics.php with other node ids.
Prints only.
"""
import json
import re
import sys

import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
SKIP = re.compile(r"google|doubleclick|facebook|analytics|cookie|clarity|hotjar|demdex|dynatrace|visualstudio|office\.net|monitor\.azure|"
                  r"approvedResources|translations|\.svg|\.woff|fonts|i18n|ui-config|openid|certs|user/self|maintenance|lotties|\.png|\.jpg", re.I)


def attributes():
    r = requests.get("https://platform.aggm.at/vis-service/api/ts/attributes", timeout=60, headers={"User-Agent": UA})
    print("AGGM attributes", r.status_code, len(r.text))
    try:
        a = r.json()["attributes"]
        for k, v in a.items():
            print(f"  {k}: {v}"[:600])
        print("  other keys:", [k for k in r.json() if k != "attributes"])
    except Exception as e:  # noqa: BLE001
        print(r.text[:500], e)
    for path in ("ts", "ts/timeseries", "ts/data", "ts/series", "ts/search", "ts/publication", "ts/list"):
        for m in ("GET",):
            try:
                rr = requests.request(m, f"https://platform.aggm.at/vis-service/api/{path}", timeout=30, headers={"User-Agent": UA})
                print(f"  {m} /{path}: {rr.status_code} {rr.text[:160]!r}")
            except Exception as e:  # noqa: BLE001
                print(f"  {path}: {type(e).__name__}")


def transgaz():
    for nid in (113, 114, 115, 112, 111, 116, 117, 120, 150):
        try:
            r = requests.post("https://www.transgaz.ro/core/modules/statistics/statistics.php", data={"nid": nid}, timeout=30,
                              headers={"User-Agent": UA, "X-Requested-With": "XMLHttpRequest", "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"})
            print(f"  Transgaz nid={nid}: {r.status_code} len={len(r.text)} {r.text[:200]!r}")
        except Exception as e:  # noqa: BLE001
            print(f"  nid={nid}: {type(e).__name__}")


PAGES = [
    ("Enagas energy data", "https://www.enagas.es/en/technical-management-system/energy-data/", ""),
    ("Enagas demand", "https://www.enagas.es/en/technical-management-system/energy-data/demand/", ""),
    ("Enagas demand history", "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/", ""),
    ("Enagas publications", "https://www.enagas.es/en/technical-management-system/energy-data/publications/", ""),
    ("Enagas bulletin", "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/", ""),
    ("Enagas informe demanda", "https://www.enagas.es/es/gestion-tecnica-sistema/energy-data/publicaciones/informe-demanda-gas-natural/", ""),
    ("Enagas boletin ES", "https://www.enagas.es/es/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico/", ""),
    ("Transgaz physical flows", "https://www.transgaz.ro/en/clients/operational-data/physical-flows", "wait"),
    ("Transgaz tech consumption", "https://www.transgaz.ro/en/clients/operational-data/technological-consumption", ""),
    ("Plinacro transparency", "https://www.plinacro.hr/default.aspx?id=109", ""),
    ("Plinacro forecast 1186", "https://www.plinacro.hr/default.aspx?id=1186", ""),
    ("Gasgrid flowing", "https://gasgrid.fi/en/gas-business/what-is-flowing-in-our-pipelines/", "wait"),
    ("Gasgrid market data", "https://gasgrid.fi/en/gas-business/market-data/", ""),
    ("Bulgartransgaz operational", "https://bulgartransgaz.bg/en/pages/operational-data-185.html", ""),
    ("Bulgartransgaz daily", "https://bulgartransgaz.bg/index/daily-reports/lang/en", ""),
    ("Gaz-System MIR", "https://swi.gaz-system.pl/mir/", "wait"),
]


def gazsystem():
    base = "https://swi.gaz-system.pl/mir/api/v1/en/"
    for m in ("", "Menu", "menu", "Consumption", "PhysicalFlows", "Flows", "NTSFlows", "ActualFlows", "Demand", "GasConsumption", "Gcv", "GasQuality"):
        try:
            r = requests.get(base + m + ("/init" if m and m.lower() != "menu" else ""), timeout=30, headers={"User-Agent": UA})
            print(f"  GS {m!r}: {r.status_code} {r.text[:300]!r}")
        except Exception as e:  # noqa: BLE001
            print(f"  GS {m}: {type(e).__name__}")


def pages():
    with sync_playwright() as p:
        b = p.chromium.launch()
        for label, url, mode in PAGES:
            ctx = b.new_context(user_agent=UA)
            pg = ctx.new_page()
            seen = []

            def on(r, seen=seen):
                if r.request.resource_type in ("xhr", "fetch", "document") and not SKIP.search(r.url):
                    try:
                        body = r.text()[:300].replace("\n", " ")
                    except Exception:  # noqa: BLE001
                        body = ""
                    seen.append(f"[{r.status}] {r.request.method} {r.url[:200]}\n        post={(r.request.post_data or '')[:400]}\n        resp={body}")
            pg.on("response", on)
            print("=" * 100, f"\n{label}: {url}", flush=True)
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=45000)
                pg.wait_for_timeout(4000)
                if mode == "click":
                    for sel in ("text=Accept", "text=I accept", "text=Agree", "text=Continue", "button", "input[type=submit]", "a.btn"):
                        try:
                            el = pg.locator(sel).first
                            if el.count():
                                print("  clicking", sel, "->", (el.inner_text() or "")[:40])
                                el.click(timeout=3000)
                                pg.wait_for_timeout(6000)
                                break
                        except Exception as e:  # noqa: BLE001
                            print("  click fail", sel, type(e).__name__)
                pg.wait_for_timeout(22000 if mode == "wait" else 6000)
            except Exception as e:  # noqa: BLE001
                print("  load problem", type(e).__name__)
            print("  title:", pg.title()[:90], "| url:", pg.url[:140])
            for s_ in seen[:16]:
                print("  req", s_)
            n = 0
            for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,70), e.href])"):
                if re.search(r"\.(xlsx?|csv|zip|json|pdf)(\?|$)|consum|demand|daily|flow|download|export|histor|data|boletin|bulletin|publication|stat", h + " " + t, re.I):
                    print(f"  link {t!r} -> {h[:180]}")
                    n += 1
                    if n >= 24:
                        break
            try:
                print("  text:", pg.inner_text("body")[:500].replace("\n", " | "))
            except Exception:  # noqa: BLE001
                pass
            ctx.close()
        b.close()


print("=" * 100)
for fn in (gazsystem, transgaz, pages):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
