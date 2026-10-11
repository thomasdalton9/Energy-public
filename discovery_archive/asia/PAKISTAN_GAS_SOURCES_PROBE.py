"""
Pakistan natural gas, fresher-source discovery (run from GitHub Actions; the sandbox is blocked).
asia/PAKISTAN_PBS_GAS.py reads the PBS Monthly Bulletin of Statistics (MBS) table 3.2 (production by field) and the PBS
monthly imports statement (LNG value only); its newest MBS data month is Mar-2024. This probe looks for monthly
production (field / province / company), consumption by sector and LNG import volumes:
  PBS media library (newest MBS / bulletin uploads), OGRA, Petroleum Division, HDIP (Energy Year Book), PPIS / DGPC,
  SNGPL / SSGC, PLL / PSO / LNG terminals, Port Qasim, SBP import payments, OGDCL / PPL / MPCL, Finance (Economic Survey).
requests first; pages that fail (403 / JS) are retried with Playwright (Chromium). Prints status, DNS, title,
gas keywords and data-like links. Nothing is written.

    python3 discovery_archive/asia/PAKISTAN_GAS_SOURCES_PROBE.py [round]
"""
import io
import re
import socket
import sys
from urllib.parse import urljoin, urlparse

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
     "Accept-Language": "en-GB,en;q=0.9"}
T = (15, 60)
KW = re.compile(r"(natural gas|MMCFD|MMSCFD|mmcft|BCF|RLNG|LNG|cargo|gas sales|consumption|year ?book|"
                r"state of (the )?(regulated )?petroleum)", re.I)
LINK = re.compile(r"""(?:href|src)=["']([^"']+)["']""", re.I)
DATA = re.compile(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)|statist|bulletin|data|download|gas|lng|report|year|book|"
                  r"publication|production|sales|annual|cargo|vessel|shipping", re.I)
FAILED = []


def out(s=""):
    print(s, flush=True)


def dns(host):
    try:
        return socket.gethostbyname(host)
    except Exception as e:  # noqa: BLE001
        return f"NO DNS ({type(e).__name__})"


def show_html(txt, base, kw=8, links=60, pat=DATA, show=0):
    t = re.search(r"<title>(.*?)</title>", txt, re.S | re.I)
    out(f"  TITLE: {re.sub(r'\s+', ' ', t.group(1)).strip()[:150] if t else None}")
    if show:
        out("  HEAD: " + re.sub(r"\s+", " ", txt[:show]))
    plain = re.sub(r"<script.*?</script>|<style.*?</style>", " ", txt, flags=re.S | re.I)
    plain = re.sub(r"<[^>]+>", " ", plain)
    plain = re.sub(r"\s+", " ", plain)
    for n, m in enumerate(KW.finditer(plain)):
        if n >= kw:
            break
        out(f"  KW: ...{plain[max(0, m.start() - 120):m.end() + 160]}...")
    ls = []
    for l in LINK.findall(txt):
        if pat.search(l) and not l.startswith(("javascript", "#", "mailto", "tel")):
            a = urljoin(base, l)
            if a not in ls and not re.search(r"\.(css|js|png|jpe?g|svg|gif|woff2?|ico)(\?|$)", a, re.I):
                ls.append(a)
    out(f"  LINKS ({len(ls)}):")
    for l in ls[:links]:
        out(f"    {l}")
    return ls


def probe(name, url, kw=8, links=60, pat=DATA, show=0, **kwa):
    host = urlparse(url).hostname
    out(f"\n{'=' * 100}\n{name}: {url}   [dns {dns(host)}]")
    try:
        r = requests.get(url, headers=H, timeout=T, verify=False, **kwa)
    except Exception as e:  # noqa: BLE001
        out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")
        FAILED.append((name, url))
        return None, []
    out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b final={r.url} "
        f"server={r.headers.get('server')} LM={r.headers.get('last-modified')}")
    if r.status_code >= 400:
        out("  BODY: " + re.sub(r"\s+", " ", r.text[:300]))
        FAILED.append((name, url))
        return r, []
    ct = r.headers.get("content-type", "")
    ls = []
    if any(x in ct for x in ("html", "json", "text", "xml")):
        ls = show_html(r.text, r.url, kw, links, pat, show)
    return r, ls


def pdf_peek(name, url, pages=4, chars=1500):
    out(f"\n{'-' * 100}\nPDF {name}: {url}")
    try:
        r = requests.get(url, headers=H, timeout=(15, 180), verify=False)
        out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b LM={r.headers.get('last-modified')}")
        if r.status_code != 200 or not r.content.startswith(b"%PDF"):
            return
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            out(f"  pages={len(pdf.pages)}")
            hits = 0
            for i, p in enumerate(pdf.pages[:200]):
                tx = p.extract_text() or ""
                if i < 1 or (re.search(r"natural gas|MMCF|RLNG|LNG", tx, re.I) and hits < pages):
                    if i >= 1:
                        hits += 1
                    out(f"  --- page {i + 1}:\n" + tx[:chars])
    except Exception as e:  # noqa: BLE001
        out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")


def wp_media(site, search, n=15):
    """Newest uploads in a WordPress media library matching `search`."""
    url = f"{site}/wp-json/wp/v2/media"
    try:
        r = requests.get(url, headers=H, timeout=T, verify=False,
                         params={"search": search, "per_page": n, "orderby": "date", "order": "desc",
                                 "_fields": "date,source_url"})
        js = r.json() if r.status_code == 200 else []
        out(f"\nWP media {site} '{search}': {r.status_code} total={r.headers.get('X-WP-Total')}")
        for m in js:
            out(f"    {m['date'][:10]}  {m['source_url']}")
        return js
    except Exception as e:  # noqa: BLE001
        out(f"\nWP media {site} '{search}': ERROR {type(e).__name__}: {str(e)[:150]}")
        return []


def browser(urls):
    """Retry failed pages with Playwright Chromium."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        out("playwright not installed")
        return
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=H["User-Agent"], ignore_https_errors=True)
        pg = ctx.new_page()
        for name, url in urls:
            out(f"\n{'=' * 100}\nBROWSER {name}: {url}")
            try:
                resp = pg.goto(url, timeout=45000, wait_until="domcontentloaded")
                pg.wait_for_timeout(4000)
                out(f"  -> {resp.status if resp else None} final={pg.url}")
                show_html(pg.content(), pg.url)
            except Exception as e:  # noqa: BLE001
                out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")
        b.close()


def round1():
    # ---- PBS: is there anything newer than the Sep-2024 MBS?
    for s in ("MBS", "Bulletin", "Monthly_Bulletin", "MSB", "Natural_Gas", "Energy", "Import_", "Statistical_Year"):
        wp_media("https://www.pbs.gov.pk", s)
    probe("PBS MBS page", "https://www.pbs.gov.pk/monthly-bulletin-of-statistics/", links=40,
          pat=re.compile(r"MBS|bulletin|\.pdf|\.xls", re.I))
    probe("PBS MBS page 2", "https://www.pbs.gov.pk/content/monthly-bulletin-statistics", links=40,
          pat=re.compile(r"MBS|bulletin|\.pdf|\.xls", re.I))
    # ---- OGRA
    for u in ("https://www.ogra.org.pk/", "https://ogra.org.pk/", "http://www.ogra.org.pk/",
              "https://www.ogra.org.pk/annual-reports", "https://www.ogra.org.pk/state-of-industry-reports",
              "https://www.ogra.org.pk/rlng-prices", "https://www.ogra.org.pk/publications"):
        probe("OGRA", u)
    # ---- Petroleum Division / Ministry of Energy
    for u in ("https://petroleum.gov.pk/", "https://www.petroleum.gov.pk/", "https://mpnr.gov.pk/",
              "https://www.mpnr.gov.pk/", "https://mope.gov.pk/", "https://petroleum.gov.pk/Detail/Publications",
              "https://www.petroleum.gov.pk/SiteImage/Downloads/"):
        probe("PetDiv", u)
    # ---- HDIP (Pakistan Energy Year Book), PPIS / DGPC
    for u in ("https://hdip.com.pk/", "https://www.hdip.com.pk/", "https://hdip.com.pk/publications",
              "https://ppis.gov.pk/", "https://www.ppis.gov.pk/", "https://ppisonline.com/", "https://www.ppisonline.com/",
              "https://dgpc.gov.pk/", "https://www.dgpc.gov.pk/", "https://ppepca.com/", "https://www.ppepca.com/"):
        probe("HDIP/PPIS/DGPC", u)
    # ---- utilities
    for u in ("https://www.sngpl.com.pk/", "https://sngpl.com.pk/", "https://www.sngpl.com.pk/web/",
              "https://www.ssgc.com.pk/", "https://www.ssgc.com.pk/web/", "https://www.isgs.com.pk/"):
        probe("SNGPL/SSGC/ISGS", u)
    # ---- LNG
    for u in ("https://pll.com.pk/", "https://www.pll.com.pk/", "https://paklng.com/", "https://www.paklng.com/",
              "https://pakistanlng.com/", "https://psopk.com/en/", "https://psopk.com/en/product-and-services/lng",
              "https://www.elengy.com/", "https://elengy.com/", "https://www.pgpcl.com.pk/", "https://pgpc.com.pk/",
              "https://www.pqa.gov.pk/", "https://pqa.gov.pk/en/shipping-intelligence", "https://www.pqa.gov.pk/en/vessels"):
        probe("LNG", u)
    # ---- SBP
    for u in ("https://www.sbp.org.pk/ecodata/index2.asp", "https://www.sbp.org.pk/ecodata/index.asp",
              "https://easydata.sbp.org.pk/"):
        probe("SBP", u, links=150, pat=re.compile(r"import|trade|commod|\.xls|\.pdf|petrol", re.I))
    # ---- E&P companies
    for u in ("https://ogdcl.com/", "https://www.ogdcl.com/", "https://www.ppl.com.pk/", "https://ppl.com.pk/",
              "https://www.mpcl.com.pk/", "https://mari.com.pk/", "https://www.pol.com.pk/"):
        probe("E&P", u)
    # ---- Finance Division (Economic Survey energy chapter, annual)
    for u in ("https://www.finance.gov.pk/survey_2526.html", "https://www.finance.gov.pk/survey_2425.html",
              "https://www.finance.gov.pk/economic_survey.html"):
        probe("Finance", u, links=40, pat=re.compile(r"survey|\.pdf", re.I))
    out("\n\nFAILED (requests): " + str(len(FAILED)))
    for n, u in FAILED:
        out(f"  {n}: {u}")
    retry = [(n, u) for n, u in FAILED if "NO DNS" not in dns(urlparse(u).hostname)]
    browser(retry)


def cf_page(ctx, url, wait=25):
    """Open url in a fresh page, wait for a Cloudflare challenge to clear; return (status, html, page)."""
    pg = ctx.new_page()
    try:
        resp = pg.goto(url, timeout=60000, wait_until="domcontentloaded")
        st = resp.status if resp else None
        for _ in range(wait):
            t = pg.title()
            if "just a moment" not in t.lower() and "security verification" not in t.lower():
                break
            pg.wait_for_timeout(1000)
        pg.wait_for_timeout(1500)
        return st, pg.content(), pg
    except Exception as e:  # noqa: BLE001
        out(f"  ERROR {type(e).__name__}: {str(e)[:160]}")
        return None, "", pg


def ctx_pdf(ctx, name, url, pages=3, chars=1800, pat=r"natural gas|MMCF|MMBTU|RLNG|LNG|cargo"):
    out(f"\n{'-' * 100}\nPDF {name}: {url}")
    try:
        r = ctx.request.get(url, timeout=180000)
        body = r.body()
        out(f"  -> {r.status} {r.headers.get('content-type')} {len(body)}b LM={r.headers.get('last-modified')}")
        if not body.startswith(b"%PDF"):
            out("  not a PDF: " + re.sub(r"\s+", " ", body[:200].decode("latin1")))
            return
        import pdfplumber
        with pdfplumber.open(io.BytesIO(body)) as pdf:
            out(f"  pages={len(pdf.pages)}")
            hits = 0
            for i, p in enumerate(pdf.pages[:250]):
                tx = p.extract_text() or ""
                if i == 0 or (re.search(pat, tx, re.I) and hits < pages):
                    hits += i > 0
                    out(f"  --- page {i + 1} ({len(tx)} chars):\n" + tx[:chars])
    except Exception as e:  # noqa: BLE001
        out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")


def round2():
    # PBS: every MBS upload (newest issue?), production / trade pages
    js = []
    for s in ("MBS", "Bulletin_of_Stat"):
        for pgn in (1, 2, 3):
            try:
                r = requests.get("https://www.pbs.gov.pk/wp-json/wp/v2/media", headers=H, timeout=T, verify=False,
                                 params={"search": s, "per_page": 100, "page": pgn, "_fields": "date,source_url"})
                if r.status_code != 200:
                    break
                js += r.json()
            except Exception:  # noqa: BLE001
                break
    names = sorted({m["source_url"].rsplit("/", 1)[1] for m in js if m["source_url"].lower().endswith(".pdf")})
    out(f"PBS MBS-like PDFs ({len(names)}): " + ", ".join(n for n in names if re.search(r"202[4-6]", n)))
    for u in ("https://www.pbs.gov.pk/pakistan-bureau-of-statistics-economic-statistics-production/",
              "https://www.pbs.gov.pk/publication-2", "https://www.pbs.gov.pk/external-trade-statistics/"):
        probe("PBS", u, links=80, kw=4, pat=re.compile(r"\.pdf|\.xls|bulletin|mbs|energy|trade|import|quantum", re.I))
    # SBP import payments / trade files
    r, ls = probe("SBP economic data", "https://www.sbp.org.pk/economic-data", links=0, kw=0)
    for l in ls:
        if re.search(r"import|export|commod|trade|petrol|oil|IP_|payment", l, re.I):
            out(f"    SBP: {l}")
    # Petroleum Division, PPIS, Mari
    for u in ("https://petroleum.gov.pk/Publications", "https://petroleum.gov.pk/Download",
              "https://ppisonline.com/data-review", "https://hdip.gov.pk/", "https://www.hdip.gov.pk/",
              "https://hdip.org.pk/"):
        probe("misc", u, links=80, kw=6)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
        ctx = b.new_context(user_agent=H["User-Agent"], ignore_https_errors=True, locale="en-GB")
        pdfs = {}
        pages = ["https://www.ogra.org.pk/", "https://www.ogra.org.pk/ogra-annual-reports",
                 "https://www.ogra.org.pk/other-reports", "https://www.ogra.org.pk/natural-gas-3",
                 "https://www.ogra.org.pk/gas-9", "https://www.ogra.org.pk/gas-16", "https://www.ogra.org.pk/lng-10",
                 "https://www.ogra.org.pk/lng-9", "https://www.ogra.org.pk/rlng-3",
                 "https://www.ogra.org.pk/rlng-notified-prices-3", "https://www.ogra.org.pk/rlng-notified-prices-for-pll",
                 "https://www.ogra.org.pk/finalactual-monthly-rlng-sale-price",
                 "https://www.ogra.org.pk/actual-rlng-notified-sale-price-decisions-ssgc-dated-august-2026",
                 "https://www.paklng.com/", "https://www.isgs.com.pk/", "https://psopk.com/",
                 "https://www.ogdcl.com/", "https://www.ppl.com.pk/", "https://www.pol.com.pk/",
                 "https://www.marienergies.com.pk/what-we-do/production-and-sales",
                 "https://www.finance.gov.pk/survey_2526.html", "https://www.finance.gov.pk/survey_2425.html",
                 "https://pqa.gov.pk/en"]
        for u in pages:
            out(f"\n{'=' * 100}\nBROWSER {u}")
            st, html, pg = cf_page(ctx, u)
            out(f"  -> {st} final={pg.url}")
            ls = show_html(html, pg.url, kw=5, links=70) if html else []
            pdfs[u] = [l for l in ls if l.lower().endswith((".pdf", ".xlsx", ".xls"))]
            pg.close()
        # sample documents: newest RLNG price notifications, other reports, annual report
        for u in ("https://www.ogra.org.pk/finalactual-monthly-rlng-sale-price",
                  "https://www.ogra.org.pk/rlng-notified-prices-for-pll", "https://www.ogra.org.pk/rlng-notified-prices-3",
                  "https://www.ogra.org.pk/other-reports", "https://www.ogra.org.pk/ogra-annual-reports",
                  "https://www.ogra.org.pk/natural-gas-3"):
            for l in pdfs.get(u, [])[:2]:
                ctx_pdf(ctx, u.rsplit("/", 1)[1], l)
        b.close()


if __name__ == "__main__":
    {"1": round1, "2": round2}[sys.argv[1] if len(sys.argv) > 1 else "1"]()
