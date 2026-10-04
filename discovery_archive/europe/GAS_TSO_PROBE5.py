"""
Probe 5: (1) Enagas history before 2023 (raw JSON for 2021/2022 query dates + Excel download links on the page);
(2) NET4GAS (CZ) transparency/data pages; (3) Gassco (Norway) flows and UMM pages; (4) AGGM ts-publication API calls (AT);
(5) Gasgrid 'what is flowing' (FI); (6) Gaz-System swi MIR measures (PL); (7) Transgaz statistics.php request/response (RO);
(8) Plinacro daily offtake page (HR). XHR requests are logged with method, URL, POST body and the start of the response.
Prints only.
"""
import json
import re
import sys

import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
ENAGAS = ("https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/demanda/historico/jcr:content/responsiveGrid/"
          "container_copy_19796/realdemand_copy_copy.realdemand.json")


def enagas():
    h = {"Accept": "application/json", "User-Agent": UA, "X-Requested-With": "XMLHttpRequest",
         "Referer": "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/"}
    for q in ("30/06/2022", "31/12/2021", "30/06/2023", "31/03/2023"):
        r = requests.get(ENAGAS, params={"date": q}, headers=h, timeout=60)
        try:
            j = r.json()
            a = j.get("actual", [])
            print(f" query {q}: {r.status_code} keys={list(j)[:8]} actual={len(a)} first={a[0] if a else None} last={a[-1] if a else None}")
        except Exception as e:  # noqa: BLE001
            print(f" query {q}: {r.status_code} {type(e).__name__} {r.text[:200]}")


SKIP = re.compile(r"google|doubleclick|facebook|analytics|cookie|clarity|hotjar|demdex|dynatrace|visualstudio|office\.net|monitor\.azure|"
                  r"approvedResources|translations|\.svg|fonts|i18n|ui-config|openid|certs|user/self|maintenance", re.I)
PAGES = [
    ("ES Enagas demand history page (links)", "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/", 0),
    ("CZ NET4GAS root", "https://www.net4gas.cz/en/", 0),
    ("CZ NET4GAS transparency", "https://www.net4gas.cz/en/transparency/", 0),
    ("NO Gassco flows", "https://www.gassco.no/en/our-activities/operations/", 0),
    ("NO Gassco root", "https://www.gassco.no/en/", 0),
    ("NO Gassco UMM", "https://umm.gassco.no/", 1),
    ("AT AGGM Endkundenverbrauch", "https://platform.aggm.at/portal/visualisation/ts-publication?src=map&granularity=day&Consumption=EndConsumer", 1),
    ("AT AGGM monthly consumption groups", "https://platform.aggm.at/portal/visualisation/ts-publication?fav=Qo1", 1),
    ("FI Gasgrid flowing", "https://gasgrid.fi/en/gas-business/what-is-flowing-in-our-pipelines/", 1),
    ("PL Gaz-System MIR", "https://swi.gaz-system.pl/mir/#/public/measure/ksp-daily-gcv?lang=en", 1),
    ("PL Gaz-System MIR root", "https://swi.gaz-system.pl/mir/#/public", 1),
    ("RO Transgaz physical flows", "https://www.transgaz.ro/en/clients/operational-data/physical-flows", 1),
    ("HR Plinacro daily offtake", "https://www.plinacro.hr/default.aspx?id=1186", 0),
]


def page_probe():
    with sync_playwright() as p:
        b = p.chromium.launch()
        for label, url, want_xhr in PAGES:
            ctx = b.new_context(user_agent=UA, accept_downloads=False)
            pg = ctx.new_page()
            seen = []

            def on(r, seen=seen):
                if r.request.resource_type in ("xhr", "fetch") and not SKIP.search(r.url):
                    try:
                        body = r.text()[:260].replace("\n", " ")
                    except Exception:  # noqa: BLE001
                        body = ""
                    pd_ = (r.request.post_data or "")[:300].replace("\n", " ")
                    seen.append(f"[{r.status}] {r.request.method} {r.url[:200]}\n        post={pd_}\n        resp={body}")
            pg.on("response", on)
            print("=" * 100, f"\n{label}: {url}", flush=True)
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=45000)
                pg.wait_for_timeout(9000)
            except Exception as e:  # noqa: BLE001
                print("  load problem", type(e).__name__)
            print("  title:", pg.title()[:90], "| url:", pg.url[:140])
            for s_ in seen[:12]:
                print("  xhr", s_)
            n = 0
            for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,70), e.href])"):
                if re.search(r"\.(xlsx?|csv|zip|json)(\?|$)|consum|verbrauch|daily|flow|download|export|histor|data|transparen|umm|publication", h + " " + t, re.I):
                    print(f"  link {t!r} -> {h[:170]}")
                    n += 1
                    if n >= 16:
                        break
            try:
                print("  text:", pg.inner_text("body")[:350].replace("\n", " | "))
            except Exception:  # noqa: BLE001
                pass
            ctx.close()
        b.close()


print("=" * 100, "\nES Enagas raw windows")
try:
    enagas()
except Exception as e:  # noqa: BLE001
    print("FAILED", type(e).__name__, str(e)[:200])
page_probe()
sys.exit(0)
