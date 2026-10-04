"""Probe: Gassco real-time feed, gassco.eu REST index, umm.gassco.no FLOW app (XHR capture, history), ENTSOG Gassco points,
Sokkeldirektoratet. Prints only."""
import json, re, sys
import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}


def get(u, n=600, **kw):
    try:
        r = requests.get(u, timeout=40, headers=H, **kw)
        print(f"GET {u} -> {r.status_code} {r.headers.get('content-type')} len={len(r.text)}\n   {r.text[:n]!r}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"GET {u} -> {type(e).__name__} {str(e)[:100]}")


def rest():
    r = get("https://gassco.eu/wp-json/gassco/v1/realTimeGraphData", 800)
    r = get("https://gassco.eu/wp-json/", 100)
    if r is not None and r.ok:
        try:
            for k in r.json().get("routes", {}):
                if "gassco" in k.lower() or "wp/v2/pages" in k:
                    print("  route", k)
        except Exception as e:  # noqa: BLE001
            print(e)
    for p in ("gassco/v1", "gassco/v1/flowData", "gassco/v1/historicalData", "gassco/v1/graphData"):
        get("https://gassco.eu/wp-json/" + p, 300)
    for u in ("https://gassco.eu/en/", "https://gassco.eu/en/data-and-statistics/", "https://gassco.eu/en/our-operations/"):
        r = get(u, 100)
        if r is not None and r.ok:
            for m in sorted(set(re.findall(r'(?:href|src)="([^"]+)"', r.text))):
                if re.search(r"xls|csv|json|statist|histor|flow|data|umm|download|wp-json|\.js", m, re.I) and "wp-includes" not in m:
                    print("   link", m[:200])


def entsog():
    base = "https://transparency.entsog.eu/api/v1/"
    r = get(base + "operatorpointdirections?limit=-1&operatorKey=NO-TSO-0001", 200)
    for q in ("operatorpointdirections?limit=-1&tSOCountry=NO", "operators?limit=-1&operatorCountryCode=NO"):
        r = get(base + q, 200)
        if r is not None and r.ok:
            try:
                j = r.json()
                key = [k for k in j if isinstance(j[k], list)][0]
                rows = j[key]
                print("  n", len(rows))
                for x in rows[:60]:
                    print("   ", {k: x[k] for k in x if k in ("operatorKey", "operatorLabel", "pointKey", "pointLabel", "directionKey", "tSOCountry", "operatorCountryCode", "adjacentSystemLabel", "adjacentOperatorKey")})
            except Exception as e:  # noqa: BLE001
                print(e)
    r = get(base + "operationalData?limit=5&indicator=Physical%20Flow&periodType=day&from=2026-09-20&to=2026-09-25&operatorKey=UK-TSO-0001&pointKey=&timezone=CET", 300)
    r = get(base + "operationalData?limit=5&indicator=Physical%20Flow&periodType=day&from=2026-09-20&to=2026-09-25&tSOCountry=NO", 300)


def other():
    for u in ("https://factpages.sodir.no/public?/Factpages/external/tableview/field_production_monthly",
              "https://www.sodir.no/en/facts/production-figures/",
              "https://www.sodir.no/en/whats-new/publications/production-figures/",
              "https://www.norskpetroleum.no/en/production-and-exports/exports-of-oil-and-gas/",
              "https://data.ssb.no/api/v0/en/table/",
              "https://gassco.no/en/our-business/",
              "https://umm.gassco.no/api/",
              ):
        r = get(u, 150)
        if r is not None and r.ok and "html" in (r.headers.get("content-type") or ""):
            for m in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
                if re.search(r"xls|csv|json|export|gas|histor|download|flow", m, re.I):
                    print("   link", m[:200])


def flow():
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=UA)
        pg = ctx.new_page()
        seen = []

        def on(r):
            if r.request.resource_type in ("xhr", "fetch", "document", "script") and not re.search(r"google|analytics|\.woff|\.css", r.url):
                try:
                    body = r.text()[:500].replace("\n", " ") if r.request.resource_type != "script" else ""
                except Exception:  # noqa: BLE001
                    body = ""
                seen.append(f"[{r.status}] {r.request.method} {r.url[:220]} post={(r.request.post_data or '')[:200]}\n      {body}")
        pg.on("response", on)
        for u in ("https://umm.gassco.no/", "https://gassco.eu/en/"):
            print("=" * 80, u)
            try:
                pg.goto(u, wait_until="domcontentloaded", timeout=45000)
                pg.wait_for_timeout(4000)
                print(" url", pg.url, "title", pg.title())
                for t in ("Accept", "I accept", "Agree", "Continue", "OK"):
                    el = pg.get_by_text(t, exact=False).first
                    if el.count():
                        print(" click", t)
                        el.click(timeout=4000)
                        pg.wait_for_timeout(8000)
                        break
                print(" url", pg.url)
                print(" text:", pg.inner_text("body")[:1500].replace("\n", " | "))
                for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,60), e.href])")[:80]:
                    print("  a", t, h[:160])
            except Exception as e:  # noqa: BLE001
                print(" problem", type(e).__name__, str(e)[:200])
            for s in seen[:60]:
                print(" req", s)
            seen.clear()
        b.close()


for fn in (rest, entsog, other, flow):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
