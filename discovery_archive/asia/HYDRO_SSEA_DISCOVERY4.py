"""
South & Southeast Asia hydro reservoir discovery, round 4 (after HYDRO_SSEA_DISCOVERY3.py):
  Vietnam   vndms.gov.vn bundles: hydro-reservoir ('Ho thuy dien') layer endpoints; try them
  Pakistan  Wayback Machine CDX coverage of IRSA's daily PDFs (pakirsa.gov.pk/Doc/Data*.pdf) to fill Dec 2024 - Sep 2026;
            PMD FFD behind a Cloudflare challenge: try cloudscraper and a headless Chromium (playwright)
"""
import re
import subprocess
import sys
from collections import Counter

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "text/html,application/json,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}
T = (15, 90)
s = requests.Session()
s.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    try:
        r = s.get(u, timeout=k.pop("timeout", T), verify=False, **k)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {type(e).__name__}: {str(e)[:200]}")
        return None


def vndms():
    out("\n=========== VNDMS")
    r = get("https://vndms.gov.vn/")
    h = r.text
    i = h.find("Hồ thủy điện")
    for m in list(re.finditer(r'(?i)(data-layer|data-url|layer[A-Za-z]*\s*[:=]|id="[^"]*(?:ho|lake|reservoir)[^"]*")', h))[:30]:
        out("  html: " + h[max(0, m.start() - 100):m.start() + 200].replace("\n", " "))
    for js in re.findall(r'<script[^>]+src=["\'](/bundles/[^"\']+)', h):
        k = get("https://vndms.gov.vn" + js)
        if k is None or not k.ok:
            continue
        t = k.text
        urls = sorted(set(re.findall(r'["\'`]((?:https?://[\w.\-]+)?/(?:api|Api|API|home|Home|map|Map|data|Data|layer|Layer|geoserver|services|ws)[\w/\-.{}?=&]*)["\'`]', t)))
        out(f"   urls in {js} ({len(urls)}):", urls[:200])
        for m in list(re.finditer(r"(?i)(hochua|ho_chua|hothuydien|thuydien|reservoir|lake)", t))[:25]:
            out("    js: " + t[max(0, m.start() - 200):m.start() + 250].replace("\n", " "))


def wayback():
    out("\n=========== Wayback IRSA")
    for pat in ("pakirsa.gov.pk/Doc/Data*", "pakirsa.gov.pk/doc/data*"):
        r = get("http://web.archive.org/cdx/search/cdx", params={"url": pat, "output": "json", "fl": "timestamp,original,statuscode",
                                                                 "collapse": "original", "limit": 5000}, timeout=(20, 180))
        if r is not None and r.ok:
            try:
                j = r.json()[1:]
            except ValueError:
                out(r.text[:500])
                continue
            out(f"  {pat}: {len(j)} captures")
            yrs = Counter()
            for ts, orig, st in j:
                m = re.search(r"Data(\d\d)-(\d\d)-(\d{4})", orig)
                if m:
                    yrs[f"{m.group(3)}-{m.group(2)}"] += 1
            out("  by month:", sorted(yrs.items()))
            out("  sample:", j[:5], j[-5:])


def ffd():
    out("\n=========== FFD")
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "cloudscraper", "playwright"], check=False)
    try:
        import cloudscraper
        sc = cloudscraper.create_scraper()
        for u in ("https://ffd.pmd.gov.pk/river-state", "https://pmd.gov.pk/FFD/index_files/daily/damssep26_files/sheet001.htm"):
            try:
                r = sc.get(u, timeout=60)
                out(f"  cloudscraper {u}: {r.status_code} {len(r.text)} {re.sub(r'<[^>]+>', ' ', r.text)[:300]}")
            except Exception as e:  # noqa: BLE001
                out("  cloudscraper err", e)
    except Exception as e:  # noqa: BLE001
        out("  cloudscraper import", e)
    subprocess.run([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"], check=False,
                   stdout=subprocess.DEVNULL)
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True)
            pg = b.new_page(user_agent=H["User-Agent"])
            for u in ("https://ffd.pmd.gov.pk/river-state", "https://pmd.gov.pk/FFD/index_files/daily/damssep26_files/sheet001.htm"):
                calls = []
                pg.on("request", lambda q: calls.append(q.url) if q.resource_type in ("xhr", "fetch") else None)
                try:
                    pg.goto(u, timeout=60000)
                    pg.wait_for_timeout(15000)
                    t = pg.inner_text("body")
                    out(f"  playwright {u}: title {pg.title()!r}; text {t[:2500]!r}")
                    out("  xhr:", calls[:40])
                except Exception as e:  # noqa: BLE001
                    out("  playwright err", u, e)
            b.close()
    except Exception as e:  # noqa: BLE001
        out("  playwright", e)


for f in (vndms, wayback, ffd):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
