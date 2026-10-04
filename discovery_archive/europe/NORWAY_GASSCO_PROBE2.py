"""Probe 2: ENTSOG physical flow at Gassco points (history/values), gassco.eu deliveryNumbers endpoints, norskpetroleum xlsx,
umm.gassco.no disclaimer form + pages after accepting. Prints only."""
import io, json, re
import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
E = "https://transparency.entsog.eu/api/v1/"


def entsog():
    for fr, to in (("2026-09-25", "2026-09-28"), ("2022-01-01", "2022-01-03"), ("2017-01-01", "2017-01-03")):
        r = requests.get(E + "operationalData", params={"limit": -1, "indicator": "Physical Flow", "periodType": "day", "from": fr, "to": to,
                                                       "operatorKey": "NO-TSO-0001", "timezone": "CET"}, timeout=120, headers=H)
        print("ENTSOG NO-TSO-0001", fr, r.status_code, len(r.text))
        try:
            rows = r.json().get("operationalData", [])
            print("  rows", len(rows))
            for x in rows[:40]:
                print("   ", {k: x.get(k) for k in ("periodFrom", "pointLabel", "directionKey", "adjacentSystemsLabel", "value", "unit", "flowStatus")})
        except Exception as e:  # noqa: BLE001
            print(r.text[:300], e)
    for ind in ("Nomination", "Allocation", "Renomination"):
        r = requests.get(E + "operationalData", params={"limit": 10, "indicator": ind, "periodType": "day", "from": "2026-09-25", "to": "2026-09-27",
                                                       "operatorKey": "NO-TSO-0001"}, timeout=120, headers=H)
        try:
            rows = r.json().get("operationalData", [])
        except Exception:  # noqa: BLE001
            rows = []
        print(ind, r.status_code, len(rows), [(x["pointLabel"], x["directionKey"], x["value"]) for x in rows[:6]])
    # UK side of the same points
    for pk in ("ITP-00022", "ITP-00091"):
        r = requests.get(E + "operationalData", params={"limit": -1, "indicator": "Physical Flow", "periodType": "day", "from": "2026-09-25", "to": "2026-09-28",
                                                       "pointKey": pk}, timeout=120, headers=H)
        try:
            rows = r.json().get("operationalData", [])
        except Exception:  # noqa: BLE001
            rows = []
        print(pk, r.status_code, len(rows), [(x["periodFrom"][:10], x["operatorKey"], x["directionKey"], x["value"]) for x in rows[:10]])


def gassco_rest():
    for p in ("deliveryNumbers", "deliveryNumbersSetup"):
        for q in ("", "?from=2026-09-01&to=2026-09-30"):
            u = f"https://gassco.eu/wp-json/gassco/v1/{p}{q}"
            try:
                r = requests.get(u, timeout=40, headers=H)
                print("GET", u, r.status_code, len(r.text), repr(r.text[:1500]))
            except Exception as e:  # noqa: BLE001
                print(u, type(e).__name__)
    r = requests.get("https://gassco.eu/wp-json/gassco/v1", timeout=40, headers=H)
    print(r.text[:1500])
    for u in ("https://gassco.eu/en/shippers/gassco-main-data-collection-gmdc/", "https://gassco.eu/en/home/"):
        r = requests.get(u, timeout=40, headers=H)
        print(u, r.status_code, len(r.text))
        for m in sorted(set(re.findall(r'(?:href|src|data-[a-z-]+)="([^"]+)"', r.text))):
            if re.search(r"xls|csv|json|histor|flow|data|umm|download|delivery|statist|rest|api", m, re.I) and "wp-includes" not in m and "wp-content/cache" not in m:
                print("   ", m[:200])
        for m in re.findall(r"(?:fetch|axios|ajax)[^;]{0,200}", r.text)[:10]:
            print("   js", m[:200])
    r = requests.get("https://gassco.eu/wp-content/cache/min/1/wp-content/themes/gassco/gutenberg/acf-blocks/gas-overview/gas-overview.js", timeout=40, headers=H)
    print("gas-overview.js", r.status_code, len(r.text))
    for m in re.findall(r".{0,150}(?:wp-json|fetch|deliveryNumbers).{0,250}", r.text)[:10]:
        print("   js", m)


def norskpetroleum():
    for u in ("https://www.norskpetroleum.no/wp-content/uploads/42-Norsk-naturgasseksport-fordelt-pa-leveransepunkt-03032026.xlsx",):
        r = requests.get(u, timeout=60, headers=H)
        print(u, r.status_code, len(r.content))
        if r.ok:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
            for ws in wb:
                print("  sheet", ws.title, ws.dimensions)
                for row in list(ws.iter_rows(values_only=True))[:25]:
                    print("   ", row)


def umm():
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=UA)
        pg = ctx.new_page()
        seen = []

        def on(r):
            if r.request.resource_type in ("xhr", "fetch", "document") and "google" not in r.url:
                try:
                    body = r.text()[:400].replace("\n", " ")
                except Exception:  # noqa: BLE001
                    body = ""
                seen.append(f"[{r.status}] {r.request.method} {r.url[:200]} post={(r.request.post_data or '')[:150]}\n      {body}")
        pg.on("response", on)
        pg.goto("https://umm.gassco.no/disclaimer", wait_until="domcontentloaded", timeout=45000)
        pg.wait_for_timeout(3000)
        html = pg.content()
        i = html.find("<form")
        print("FORM HTML:", re.sub(r"\s+", " ", html[i:i + 1500]) if i >= 0 else "no form")
        print("buttons:", pg.eval_on_selector_all("button,input,a.button,a.btn", "els=>els.map(e=>[e.tagName,e.type,e.id,e.className,e.value,e.innerText.slice(0,40)])"))
        for sel in ("input[type=submit]", "button[type=submit]", "button", "input[type=button]", "a.btn", "a.button", "text=Accept", "text=accept"):
            el = pg.locator(sel).first
            if el.count():
                print("click", sel)
                try:
                    el.click(timeout=4000)
                    pg.wait_for_timeout(8000)
                    break
                except Exception as e:  # noqa: BLE001
                    print(" fail", type(e).__name__)
        print("url after:", pg.url, pg.title())
        print("text:", pg.inner_text("body")[:2500].replace("\n", " | "))
        for t, h in pg.eval_on_selector_all("a[href]", "els=>els.map(e=>[e.innerText.trim().replace(/\\s+/g,' ').slice(0,60), e.href])")[:80]:
            print("  a", t, h[:160])
        for s in seen[-30:]:
            print(" req", s)
        b.close()


for fn in (entsog, gassco_rest, norskpetroleum, umm):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
