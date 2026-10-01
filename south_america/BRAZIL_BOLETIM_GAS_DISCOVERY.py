"""
MME replaced the PDF/xlsx gas bulletin (whose historical-balance annex
froze at 2025-06) with an interactive "Boletim do Gas" on the
Observatorio de Minas e Energia:
  https://www.gov.br/mme/pt-br/assuntos/observatorio-de-minas-e-energia/petroleo-gas-e-biocombustiveis/boletim-do-gas
Find what feeds it: list iframes/links on the page, then load it in a
browser and capture every data request (Power BI publish-to-web
querydata calls, Qlik, Tableau, JSON/CSV/XLSX) with request bodies and
response samples, so a scraper can replay them. Also follow any
embedded dashboard URL directly. Not reachable from the editing sandbox.
"""
import json
import re

import requests

PAGE = "https://www.gov.br/mme/pt-br/assuntos/observatorio-de-minas-e-energia/petroleo-gas-e-biocombustiveis/boletim-do-gas"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def out(*a):
    print(*a, flush=True)


r = requests.get(PAGE, headers=HEADERS, timeout=(10, 60))
out(f"GET {PAGE} -> {r.status_code} {len(r.text)} chars")
for m in re.findall(r"<iframe[^>]*>", r.text, re.I):
    out("  iframe:", m[:400])
for m in sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', r.text))):
    if re.search(r"powerbi|qlik|tableau|arcgis|looker|datastudio|\.xlsx?|\.csv|\.json|painel|dashboard|boletim", m, re.I):
        out("  link:", m)
text = re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S | re.I)
text = re.sub(r"\s+", " ", text)
i = text.lower().find("boletim do g")
out("  text:", text[max(0, i - 200): i + 1500])

from playwright.sync_api import sync_playwright

INTERESTING = re.compile(r"querydata|conceptualschema|modelsAndExploration|/public/reports|wabi|qlik|tableau|"
                         r"\.(csv|xlsx?|json)(\?|$)|/api/", re.I)
with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(user_agent=HEADERS["User-Agent"], viewport={"width": 1600, "height": 1100})
    page = ctx.new_page()
    reqs = []
    page.on("request", lambda rq: reqs.append(rq) if INTERESTING.search(rq.url) and "analytics" not in rq.url else None)
    page.goto(PAGE, timeout=90000, wait_until="load")
    page.wait_for_timeout(25000)
    out("\ntitle:", page.title())
    for fr in page.frames:
        out("  frame:", fr.url[:300])
    embed = [fr.url for fr in page.frames if re.search(r"powerbi|qlik|tableau", fr.url, re.I)]
    if embed:
        out("\n== opening embed directly:", embed[0][:300])
        p2 = ctx.new_page()
        p2.on("request", lambda rq: reqs.append(rq) if INTERESTING.search(rq.url) else None)
        p2.goto(embed[0], timeout=90000, wait_until="load")
        p2.wait_for_timeout(25000)
        out("  embed title:", p2.title())
        for t in p2.locator("[role=tab], .pbi-glyph-chevronrightmedium, .navigation-wrapper button").all()[:20]:
            try:
                out("   nav:", (t.get_attribute("aria-label") or t.inner_text() or "")[:80])
            except Exception:
                pass
    out(f"\n{len(reqs)} interesting requests")
    seen = set()
    for rq in reqs:
        key = (rq.method, rq.url.split("?")[0])
        if key in seen and "querydata" not in rq.url.lower():
            continue
        seen.add(key)
        out(f"\n--- {rq.method} {rq.url[:300]}")
        hdr = {k: v for k, v in rq.headers.items() if k.lower() in ("x-powerbi-resourcekey", "authorization", "activityid", "content-type")}
        if hdr:
            out("   headers:", json.dumps({k: v[:80] for k, v in hdr.items()}))
        if rq.post_data:
            out("   body:", rq.post_data[:3000])
        try:
            resp = rq.response()
            if resp:
                body = resp.text()
                out(f"   resp {resp.status} {len(body)}B:", body[:3000])
        except Exception as e:
            out("   resp err", type(e).__name__)
    ctx.close()
    b.close()
