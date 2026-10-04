"""
Probe 2 for NL / IT national gas consumption: Eurostat nrg_cb_gasm (IT, NL: balance items, latest month, 2025 values),
CBS StatLine gas tables (NL), Snam 'Physical Flows' and 'Commercial Flows' pages (links, XHR), MASE monthly gas balance.
Prints only.
"""
import json
import re
import sys

import requests
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"


def eurostat():
    B = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm"
    for geo in ("IT", "NL"):
        r = requests.get(B, params={"format": "JSON", "lang": "EN", "geo": geo, "sinceTimePeriod": "2025-01"}, timeout=180)
        print(f"Eurostat {geo}: {r.status_code}", flush=True)
        if r.status_code != 200:
            print(r.text[:300])
            continue
        j = r.json()
        dims, size = j["id"], j["size"]
        print(" dims", dims, size)
        for d in dims:
            if d in ("time", "geo", "freq"):
                continue
            print("  ", d, json.dumps(j["dimension"][d]["category"].get("label", {}))[:900])
        vals = j["value"]
        pos = lambda c: sum(c[d] * __import__("math").prod(size[i + 1:]) for i, d in enumerate(dims))
        idx = {d: j["dimension"][d]["category"]["index"] for d in dims}
        times = idx["time"]
        for b, bi in idx["nrg_bal"].items():
            for u, ui in idx.get("unit", {"": 0}).items():
                row = []
                for t, ti in times.items():
                    c = {d: 0 for d in dims}
                    c["nrg_bal"], c["time"] = bi, ti
                    if "unit" in dims:
                        c["unit"] = ui
                    v = vals.get(str(pos(c)))
                    row.append(None if v is None else round(v))
                if any(x is not None for x in row):
                    print(f"   {b} {u}: {row}")
        print(" times:", list(times))


def cbs():
    for q in ("gas", "aardgas", "natural gas"):
        r = requests.get("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$format": "json", "$filter": f"substringof('{q}',Title) eq true", "$select": "Identifier,Title,Period,Frequency"}, timeout=60, headers={"User-Agent": UA})
        print(f"CBS catalog '{q}':", r.status_code)
        if r.ok:
            for t in r.json().get("value", [])[:25]:
                print("   ", t.get("Identifier"), "|", t.get("Title"), "|", t.get("Period"), "|", t.get("Frequency"))


def snam_pages():
    urls = ["https://www.snam.it/en/our-businesses/transportation/business-information/physical-flows.html",
            "https://www.snam.it/en/our-businesses/transportation/business-information/commercial-flows.html",
            "https://www.snam.it/it/le-nostre-attivita/trasporto/informazioni-di-business/flussi-fisici.html"]
    with sync_playwright() as p:
        b = p.chromium.launch()
        for u in urls:
            ctx = b.new_context(user_agent=UA)
            pg = ctx.new_page()
            seen = []
            pg.on("response", lambda r: seen.append(f"[{r.status}] {r.url[:180]}") if r.request.resource_type in ("xhr", "fetch") and "snam" in r.url else None)
            try:
                pg.goto(u, wait_until="networkidle", timeout=45000)
            except Exception as e:  # noqa: BLE001
                print(" load problem", type(e).__name__)
            print("=" * 80, "\nSnam", u, "->", pg.url, "|", pg.title())
            for s in seen[:20]:
                print("  xhr", s)
            for t, h in pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().slice(0,70), e.href])"):
                if re.search(r"jarvis|\.xlsx?|\.csv|flows|flussi|balance|bilancio|daily|giorn", h + t, re.I):
                    print(f"  link {t!r} -> {h[:200]}")
            ctx.close()
        b.close()


def mase():
    for u in ("https://dgsaie.mise.gov.it/bilancio-gas-naturale", "https://dgsaie.mise.gov.it/importazioni-esportazioni-gas-naturale",
              "https://www.mase.gov.it/energia/statistiche-energetiche"):
        try:
            r = requests.get(u, timeout=45, headers={"User-Agent": UA})
            print("MASE", u, r.status_code, len(r.text))
            for m in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv))"', r.text)))[:15]:
                print("   ", m)
        except Exception as e:  # noqa: BLE001
            print("MASE", u, type(e).__name__, str(e)[:100])


for fn in (eurostat, cbs, mase, snam_pages):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
