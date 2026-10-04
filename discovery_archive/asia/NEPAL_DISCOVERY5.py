"""Probe 5: (a) Wayback Machine holdings of the old NEA LDC NDOR PDFs (the live /admin/assets/uploads/ldc paths now
404); (b) transd/genrd 'operational reports' pages and their menus on the new site (headless Chromium)."""
import io
import re
import json
import requests
from collections import Counter
from playwright.sync_api import sync_playwright
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
W = requests.Session()
W.headers["User-Agent"] = UA

allrows = []
for host in ("nea.org.np", "www.nea.org.np", "cms.nea.org.np"):
    for prefix in ("admin", "admin%5C", "admin\\"):
        try:
            r = W.get("https://web.archive.org/cdx/search/cdx", params={
                "url": f"{host}/{prefix}", "matchType": "prefix", "output": "json", "filter": "original:.*NDOR.*",
                "fl": "original,timestamp,statuscode,mimetype,length", "limit": 20000}, timeout=180)
            rows = r.json()[1:] if r.ok and r.text.strip() else []
            print("CDX", host, prefix, r.status_code, len(rows), flush=True)
            allrows += rows
        except Exception as e:
            print("CDX ERR", host, prefix, type(e).__name__, str(e)[:150])
ok = [x for x in allrows if x[2] == "200" and "pdf" in x[3]]
names = sorted(set(re.search(r"NDOR[ %_]*(\d{4})[_ ](\d\d)[_ ](\d\d)", x[0].replace("%20", " ")).groups()
                   for x in ok if re.search(r"NDOR[ %_]*(\d{4})[_ ](\d\d)[_ ](\d\d)", x[0].replace("%20", " "))))
print("distinct NDOR BS dates with a 200 PDF capture:", len(names))
print("by BS year-month:", sorted(Counter(f"{y}-{m}" for y, m, d in names).items()))
for x in ok[:5]:
    print("  ", x)
if ok:
    x = ok[len(ok) // 2]
    u = f"https://web.archive.org/web/{x[1]}id_/{x[0]}"
    r = W.get(u, timeout=120)
    print("WAYBACK FETCH", r.status_code, len(r.content), r.content[:5], u)
    if r.content[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as d:
            for pg in d.pages[:2]:
                print("-----"); print(pg.extract_text())

with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)

    def visit(url, wait=4000):
        page = ctx.new_page()
        try:
            r = page.goto(url, wait_until="networkidle", timeout=90000)
            page.wait_for_timeout(wait)
            html = page.content()
            print(f"\n=== {r.status if r else '?'} {page.url} len {len(html)} title {page.title()}", flush=True)
            return page, html
        except Exception as e:
            print("VISIT ERR", url, type(e).__name__, str(e)[:200])
            return page, ""

    for u in ("https://transd.nea.org.np/en/pages/transmission-operational-reports",
              "https://genrd.nea.org.np/en/pages/generation-operational-reports", "https://transd.nea.org.np/en"):
        page, html = visit(u)
        body = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>", " ", html, flags=re.S)
        i = body.find("Operational Reports")
        print("MAIN HTML:", re.sub(r"\s+", " ", body[i:i + 4000]) if i >= 0 else re.sub(r"\s+", " ", body[-4000:]))
        for l in sorted(set(re.findall(r'(?:href|src|data-src)="([^"]+)"', html))):
            if not l.startswith("data:") and ("transd" in l or "genrd" in l or "pdf" in l.lower() or "iframe" in l):
                print("   ->", l)
        for f in page.frames:
            print("   FRAME", f.url)
        page.close()
    for u in ("https://nea.org.np/en/category/other-reports", "https://nea.org.np/en/category/publication-and-reports",
              "https://nea.org.np/en/category/auto-publication", "https://nea.org.np/en/category/downloads"):
        page, html = visit(u, 2500)
        t = re.sub(r"\s+", " ", page.inner_text("body")) if html else ""
        i = t.find("Reports") if "Reports" in t else 0
        print("TEXT", t[i:i + 1500])
        for l in sorted(set(re.findall(r'href="([^"]+)"', html))):
            if re.search(r"/detail/|pdf|page=", l):
                print("   ->", l)
        page.close()
    b.close()
