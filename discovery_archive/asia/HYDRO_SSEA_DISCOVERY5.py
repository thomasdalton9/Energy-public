"""
South & Southeast Asia hydro reservoir discovery, round 5 (after HYDRO_SSEA_DISCOVERY4.py):
  Pakistan  PMD FFD via cloudscraper (passes the Cloudflare check): /river-state/data JSON, other FFD pages with dam
            levels / history (links on the site)
  Vietnam   VNDMS: ajax url strings in the bundles; headless Chromium capture of XHR calls when the hydro-reservoir
            layer is switched on
"""
import json
import re
import subprocess
import sys

subprocess.run([sys.executable, "-m", "pip", "install", "-q", "cloudscraper", "playwright"], check=False)
import cloudscraper  # noqa: E402

sc = cloudscraper.create_scraper()


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    try:
        r = sc.get(u, timeout=60, **k)
        out(f"GET {u} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {e}")
        return None


def ffd():
    out("\n=========== FFD")
    r = get("https://ffd.pmd.gov.pk/river-state/data")
    if r is not None and r.ok:
        t = r.text
        out("  data:", t[:3000])
        for m in list(re.finditer(r"(?i)tarbela|mangla", t))[:6]:
            out("   near: " + t[max(0, m.start() - 300):m.start() + 600])
    r = get("https://ffd.pmd.gov.pk/river-state")
    links = set()
    if r is not None:
        links |= set(re.findall(r'href=["\']([^"\'#]+)', r.text))
        for m in list(re.finditer(r"(fetch\(|/data|axios|\.json)", r.text))[:15]:
            out("   js: " + r.text[max(0, m.start() - 150):m.start() + 250].replace("\n", " "))
    r = get("https://ffd.pmd.gov.pk/")
    if r is not None:
        links |= set(re.findall(r'href=["\']([^"\'#]+)', r.text))
        out("  home text:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(script|style).*?</\1>", "", r.text)))[:2000])
    out("  links:", sorted(links))
    for u in sorted(links):
        if re.search(r"(?i)dam|reservoir|daily|report|tarbela|mangla|river|flow|archive|data", u) and not u.endswith((".css", ".js", ".png")):
            full = u if u.startswith("http") else "https://ffd.pmd.gov.pk" + ("" if u.startswith("/") else "/") + u
            if "pmd.gov.pk" not in full:
                continue
            p = get(full)
            if p is not None and p.ok:
                txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(script|style).*?</\1>", "", p.text)))
                out("    text:", txt[:1200])
                i = txt.lower().find("tarbela")
                if i >= 0:
                    out("    tarbela:", txt[max(0, i - 200):i + 1200])
                out("    sublinks:", sorted(set(re.findall(r'href=["\']([^"\'#]+\.(?:pdf|xlsx?|htm|csv|json))', p.text)))[:40])
    for u in ("https://ffd.pmd.gov.pk/dams", "https://ffd.pmd.gov.pk/reservoirs", "https://ffd.pmd.gov.pk/daily-report",
              "https://ffd.pmd.gov.pk/reports", "https://ffd.pmd.gov.pk/river-state/history",
              "https://ffd.pmd.gov.pk/river-state/data?date=2025-06-01", "https://ffd.pmd.gov.pk/api/river-state"):
        p = get(u)
        if p is not None and p.ok:
            out("   ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", p.text))[:600])


def vndms():
    out("\n=========== VNDMS")
    import requests
    for js in ("vndms-core-e817ad2438.min.js", "vndms-components-ca5dfe7d0b.min.js", "home-8cbdc5b369.min.js"):
        t = requests.get("https://vndms.gov.vn/bundles/" + js, timeout=60).text
        u = sorted(set(re.findall(r'["\'`](/?[A-Za-z]+/[A-Za-z][\w\-/]*)["\'`]', t)))
        out(f"  {js} path-like strings ({len(u)}):", u[:300])
        for m in list(re.finditer(r"(?i)(ajax|\$\.get|\$\.post|getJSON|fetch\()", t))[:12]:
            out("    js: " + t[max(0, m.start() - 100):m.start() + 250].replace("\n", " "))
    subprocess.run([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"], check=False,
                   stdout=subprocess.DEVNULL)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        pg = b.new_page()
        calls = []

        def on_resp(resp):
            if resp.request.resource_type in ("xhr", "fetch"):
                try:
                    body = resp.text()[:600]
                except Exception:  # noqa: BLE001
                    body = ""
                calls.append((resp.url, resp.status, body))
        pg.on("response", on_resp)
        pg.goto("https://vndms.gov.vn/", timeout=90000)
        pg.wait_for_timeout(12000)
        out("  calls after load:")
        for c in calls:
            out("   ", c[0], c[1], c[2][:300].replace("\n", " "))
        n = len(calls)
        for sel in ("text=Hồ thủy điện", "text=Hồ chứa", "text=Thủy điện"):
            try:
                els = pg.locator(sel)
                out(f"  {sel}: {els.count()} elements")
                for i in range(min(els.count(), 3)):
                    try:
                        els.nth(i).click(timeout=5000, force=True)
                        pg.wait_for_timeout(6000)
                    except Exception as e:  # noqa: BLE001
                        out("   click err", str(e)[:150])
            except Exception as e:  # noqa: BLE001
                out("  sel err", e)
        out("  calls after clicks:")
        for c in calls[n:]:
            out("   ", c[0], c[1], c[2][:600].replace("\n", " "))
        b.close()


for f in (ffd, vndms):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
