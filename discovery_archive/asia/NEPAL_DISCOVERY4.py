"""Probe 4: the new nea.org.np (Laravel, behind F5 TSPD; passes in headless Chromium). Old /admin/assets/uploads/ldc
PDFs now 404. Find where the LDC daily operation reports live: all home links, the 'Energy Details' widget, NEA
subdomains (genrd/transd have /pages/...-operational-reports), site search, wayback CDX."""
import io
import re
import json
import socket
import requests
import urllib3
from playwright.sync_api import sync_playwright
urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# 1. wayback CDX of old LDC uploads (history of file names) and of new-site report pages
for q in ("nea.org.np/admin/assets/uploads/ldc/*", "nea.org.np/storage/*NDOR*", "cms.nea.org.np/*ldc*"):
    try:
        r = requests.get("https://web.archive.org/cdx/search/cdx", params={"url": q, "limit": 40, "output": "json",
                         "collapse": "urlkey"}, timeout=60, headers={"User-Agent": UA})
        print("CDX", q, r.status_code, r.text[:3000])
    except Exception as e:
        print("CDX ERR", q, e)

# 2. subdomains
for sub in ("sod", "ldc", "soc", "sos", "sodd", "so", "nldc", "gso", "genrd", "transd", "dcsd", "pmitd", "cms",
            "admin", "api", "files", "storage", "eservice"):
    h = f"{sub}.nea.org.np"
    try:
        ip = socket.gethostbyname(h)
    except Exception:
        print("DNS no", h)
        continue
    print("DNS", h, ip)

calls = []
with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)
    page = ctx.new_page()
    page.on("response", lambda r: calls.append((r.status, r.request.resource_type, r.url))
            if r.request.resource_type in ("xhr", "fetch", "document") else None)

    def visit(url, n=1500, wait=3000):
        try:
            r = page.goto(url, wait_until="networkidle", timeout=90000)
            page.wait_for_timeout(wait)
            html = page.content()
            print(f"\n=== {r.status if r else '?'} {page.url} len {len(html)} title {page.title()}", flush=True)
            t = re.sub(r"\s+", " ", page.inner_text("body"))
            print("TEXT", t[:n])
            return html
        except Exception as e:
            print("VISIT ERR", url, type(e).__name__, str(e)[:200])
            return ""

    html = visit("https://nea.org.np/en", 200, 6000)
    hrefs = sorted(set(h for h in re.findall(r'href="([^"]+)"', html) if not h.startswith("data:")))
    print("ALL LINKS:")
    for h in hrefs:
        print("  ", h)
    i = html.find("Energy Details")
    print("ENERGY WIDGET HTML:", html[max(0, i - 500):i + 3500])
    for u in ("https://genrd.nea.org.np/pages/generation-operational-reports",
              "https://transd.nea.org.np/pages/transmission-operational-reports",
              "https://sod.nea.org.np/", "https://ldc.nea.org.np/",
              "https://nea.org.np/en/category/auto-publication", "https://nea.org.np/category/other-reports",
              "https://nea.org.np/en/category/publication-and-reports", "https://nea.org.np/category/downloads",
              "https://nea.org.np/en/search?q=NDOR", "https://nea.org.np/en/search?keyword=daily+operation"):
        h2 = visit(u, 1200)
        for l in sorted(set(re.findall(r'href="([^"]+)"', h2))):
            if re.search(r"pdf|ldc|ndor|daily|operation|report|storage|pages/|page=|category/", l, re.I) and \
                    not l.startswith("data:"):
                print("    ->", l)
    b.close()
print("\nNET:")
for s, t, u in calls:
    if "TSPD" not in u:
        print("  ", s, t, u[:220])
