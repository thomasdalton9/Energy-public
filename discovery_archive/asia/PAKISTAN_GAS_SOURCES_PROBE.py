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


def text_of(html, n=2500):
    plain = re.sub(r"<script.*?</script>|<style.*?</style>|<nav.*?</nav>|<header.*?</header>", " ", html, flags=re.S | re.I)
    plain = re.sub(r"<[^>]+>", " ", plain)
    return re.sub(r"\s+", " ", plain)[:n]


def xls_rows(name, content, pat=r"LNG|LIQUEFIED|LIQUIFIED|natural gas", n=12):
    import pandas as pd
    try:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:  # noqa: BLE001
        out(f"  {name}: read error {type(e).__name__}: {str(e)[:150]}")
        return
    for sh, df in list(sheets.items())[:6]:
        out(f"  [{name}] sheet {sh!r} shape={df.shape}")
        for i in range(min(8, len(df))):
            out("    hdr " + " | ".join(str(v)[:18] for v in df.iloc[i].tolist()[:16] if str(v) != "nan"))
        hits = [i for i in range(len(df)) if re.search(pat, " ".join(map(str, df.iloc[i, :6])), re.I)]
        for i in hits[:n]:
            out(f"    row {i}: " + " | ".join(str(v)[:14] for v in df.iloc[i].tolist()[:24]))


def round3():
    # ---- Cloudflare sites through curl_cffi (Chrome TLS fingerprint)
    cf_urls = ["https://www.ogra.org.pk/", "https://www.ogra.org.pk/other-reports", "https://www.ogra.org.pk/ogra-annual-reports",
               "https://www.ogra.org.pk/rlng-notified-prices-for-pll", "https://www.ogra.org.pk/natural-gas-3",
               "https://www.paklng.com/", "https://www.isgs.com.pk/", "https://www.ogdcl.com/", "https://www.ppl.com.pk/",
               "https://www.finance.gov.pk/survey_2425.html"]
    try:
        from curl_cffi import requests as cr
        for u in cf_urls:
            try:
                r = cr.get(u, impersonate="chrome", timeout=40)
                t = re.search(r"<title>(.*?)</title>", r.text, re.S | re.I)
                out(f"curl_cffi {u}: {r.status_code} {len(r.content)}b title={t.group(1).strip()[:80] if t else None}")
            except Exception as e:  # noqa: BLE001
                out(f"curl_cffi {u}: ERROR {type(e).__name__}: {str(e)[:120]}")
    except ImportError:
        out("curl_cffi missing")
    # ---- Wayback: which OGRA / PLL documents exist (names, dates)
    for q in ("ogra.org.pk/images/data/downloads/*", "ogra.org.pk/images/*", "ogra.org.pk/*", "paklng.com/*"):
        try:
            r = requests.get("http://web.archive.org/cdx/search/cdx", timeout=(15, 120),
                             params={"url": q, "output": "json", "filter": "original:.*\\.(pdf|xlsx?)$",
                                     "collapse": "urlkey", "limit": 4000, "fl": "timestamp,original"})
            rows = r.json()[1:] if r.status_code == 200 and r.text.strip() else []
            out(f"\nWAYBACK {q}: {r.status_code} {len(rows)} docs")
            keep = [x for x in rows if re.search(r"gas|lng|rlng|state|industry|annual|report|price|sngpl|ssgc|cargo|year", x[1], re.I)]
            for ts, o in sorted(keep, key=lambda x: x[0])[-80:]:
                out(f"    {ts[:8]} {o[:170]}")
        except Exception as e:  # noqa: BLE001
            out(f"WAYBACK {q}: ERROR {type(e).__name__}: {str(e)[:150]}")
    # ---- HDIP Pakistan Energy Yearbook
    r, ls = probe("HDIP yearbook", "https://hdip.gov.pk/pakistan-energy-yearbook", links=60, kw=6)
    if r is not None and r.status_code == 200:
        out("  TEXT: " + text_of(r.text, 2000))
        for l in [l for l in ls if l.lower().endswith(".pdf")][:2]:
            pdf_peek("HDIP", l, pages=5, chars=1600)
    # ---- PBS energy & mining page
    r, ls = probe("PBS energy-mining", "https://www.pbs.gov.pk/energy-mining-2/", links=60, kw=4,
                  pat=re.compile(r"\.pdf|\.xls", re.I))
    # ---- PQA daily shipping / berth-wise cargo (LNG terminals)
    for u in ("https://pqa.gov.pk/pqa/portoperation/DailyShippingProgram", "https://pqa.gov.pk/pqa/portoperation/BerthWiseCargoHandling",
              "https://pqa.gov.pk/en/documents/album/daily-shipping-and-cargo-handling-reports"):
        r, ls = probe("PQA", u, links=40, kw=6, pat=re.compile(r"\.pdf|\.xls|report|album|page", re.I))
        if r is not None and r.status_code == 200:
            t = text_of(r.text, 200000)
            for m in list(re.finditer(r"LNG|EETL|EETPL|PGPC|Elengy|GasPort", t))[:8]:
                out(f"  LNGCTX: ...{t[max(0, m.start() - 200):m.end() + 250]}...")
    # ---- SBP import payments by commodity (LNG row) and import-of-goods workbook
    for u in ("https://www.sbp.org.pk/assets/document/Import_Payments_by_Commodities_and_Groups_Arch.xls",
              "https://www.sbp.org.pk/assets/document/Import_Payments_by_all_Commodities_Arch.xls",
              "https://www.sbp.org.pk/assets/document/import-Jan24-Onward.xlsx"):
        try:
            r = requests.get(u, headers=H, timeout=(15, 180), verify=False)
            out(f"\nSBP {u}: {r.status_code} {r.headers.get('content-type')} {len(r.content)}b LM={r.headers.get('last-modified')}")
            if r.status_code == 200:
                xls_rows(u.rsplit("/", 1)[1], r.content)
        except Exception as e:  # noqa: BLE001
            out(f"SBP {u}: ERROR {e}")
    # ---- Mari production page, PSO LNG page (text)
    for u in ("https://www.marienergies.com.pk/what-we-do/production-and-sales",
              "https://psopk.com/en/fuels/gaseous-fuels/liquefied-natural-gas"):
        r, ls = probe("text", u, links=20, kw=0)
        if r is not None and r.status_code == 200:
            out("  TEXT: " + text_of(r.text, 3000))
    # ---- headful Chromium (xvfb) through Cloudflare: OGRA home, then follow links in the same tab
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return
    with sync_playwright() as p:
        for chan in ("chrome", None):
            try:
                b = p.chromium.launch(headless=False, channel=chan,
                                      args=["--disable-blink-features=AutomationControlled"])
            except Exception as e:  # noqa: BLE001
                out(f"launch {chan}: {e}")
                continue
            ctx = b.new_context(viewport={"width": 1366, "height": 900}, locale="en-GB")
            ctx.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            pg = ctx.new_page()
            out(f"\n#### headful channel={chan}")
            ok = 0
            for u in ("https://www.ogra.org.pk/", "https://www.ogra.org.pk/other-reports",
                      "https://www.ogra.org.pk/rlng-notified-prices-for-pll", "https://www.ogra.org.pk/natural-gas-3",
                      "https://www.ogra.org.pk/ogra-annual-reports", "https://www.paklng.com/", "https://www.isgs.com.pk/",
                      "https://www.ogdcl.com/"):
                try:
                    pg.goto(u, timeout=60000, wait_until="domcontentloaded")
                    for _ in range(40):
                        if "just a moment" not in pg.title().lower():
                            break
                        pg.wait_for_timeout(1000)
                    pg.wait_for_timeout(2000)
                    t = pg.title()
                    out(f"\nHEADFUL {u}: title={t[:90]}")
                    if "just a moment" not in t.lower():
                        ok += 1
                        ls = show_html(pg.content(), pg.url, kw=4, links=60)
                        docs = [l for l in ls if re.search(r"\.(pdf|xlsx?)$", l, re.I)]
                        for l in docs[:1]:
                            ctx_pdf(ctx, "doc", l, pages=3)
                except Exception as e:  # noqa: BLE001
                    out(f"HEADFUL {u}: ERROR {type(e).__name__}: {str(e)[:150]}")
            b.close()
            if ok > 2:
                break


def anchors(html, base, pat=r"/download/|\.pdf|\.xlsx?"):
    """(url, anchor text / table-row text) for document links."""
    res = []
    for m in re.finditer(r"<tr.*?</tr>|<li.*?</li>|<a\b[^>]*>.*?</a>", html, re.S | re.I):
        blk = m.group(0)
        for h in re.findall(r"""href=["']([^"']+)["']""", blk, re.I):
            if re.search(pat, h, re.I):
                txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", blk)).strip()
                u = urljoin(base, h)
                if u not in [x for x, _ in res]:
                    res.append((u, txt[:160]))
    return res


def cf_get(url, **kw):
    from curl_cffi import requests as cr
    for i in range(3):
        try:
            r = cr.get(url, impersonate="chrome", timeout=180, **kw)
            if r.status_code == 200:
                return r
            out(f"  cf_get {url}: {r.status_code}")
        except Exception as e:  # noqa: BLE001
            out(f"  cf_get {url}: {type(e).__name__}: {str(e)[:120]}")
    return None


def pdf_text(name, body, pat=r"natural gas|MMCF|MMBTU|RLNG|LNG|cargo|sector", pages=4, chars=2000, first=1):
    import pdfplumber
    if not body[:5].startswith(b"%PDF"):
        out(f"  {name}: not a PDF: {body[:120]!r}")
        return
    with pdfplumber.open(io.BytesIO(body)) as pdf:
        out(f"  {name}: pages={len(pdf.pages)} size={len(body)}b")
        hits = 0
        for i, p in enumerate(pdf.pages[:400]):
            tx = p.extract_text() or ""
            if i < first or (re.search(pat, tx, re.I) and hits < pages):
                hits += i >= first
                out(f"  --- page {i + 1} ({len(tx)} chars):\n" + tx[:chars])


def round4():
    # ---- OGRA documents (curl_cffi passes Cloudflare)
    lists = {}
    for slug in ("petroleum-industry-reports", "annual-reports-2", "ogra-annual-reports", "other-reports",
                 "rlng-notified-prices-for-pll", "rlng-notified-prices-3", "finalactual-monthly-rlng-sale-price",
                 "natural-gas-3", "lng-9", "lng-10", "gas-9", "gas-16", "rlng-3"):
        u = f"https://www.ogra.org.pk/{slug}"
        r = cf_get(u)
        if r is None:
            continue
        a = anchors(r.text, u)
        lists[slug] = a
        out(f"\nOGRA {slug}: {len(a)} documents")
        for x, t in a[:25]:
            out(f"    {x}  ::  {t}")
    # samples: newest petroleum industry report, newest PLL RLNG price notification, natural gas doc
    for slug, n, pat in (("petroleum-industry-reports", 1, r"natural gas|gas (supply|consumption|sales)|sector|MMCF"),
                         ("rlng-notified-prices-for-pll", 2, r"cargo|MMBTU|quantity|volume|LNG"),
                         ("finalactual-monthly-rlng-sale-price", 1, r"cargo|MMBTU|quantity|volume|LNG"),
                         ("natural-gas-3", 1, r"MMCF|sector|consumption|sales")):
        for x, t in lists.get(slug, [])[:n]:
            out(f"\n{'-' * 100}\nOGRA DOC {slug}: {x} :: {t}")
            r = cf_get(x)
            if r is not None:
                out(f"  ct={r.headers.get('content-type')} cd={r.headers.get('content-disposition')}")
                pdf_text(slug, r.content, pat=pat, pages=5)
    # ---- Mari Energies production page via curl_cffi
    r = cf_get("https://www.marienergies.com.pk/what-we-do/production-and-sales")
    if r is not None:
        out("\nMARI TEXT: " + text_of(r.text, 3000))
    # ---- PQA daily shipping & cargo report: LNG lines, history depth
    r = requests.get("https://pqa.gov.pk/en/documents/album/daily-shipping-and-cargo-handling-reports?page=245",
                     headers=H, timeout=T, verify=False)
    old = re.findall(r"dailyShippingCargoStat_[^\"']+", r.text)
    out(f"\nPQA page 245: {r.status_code} oldest files: {old[-3:]}")
    for f in ("https://www.pqa.gov.pk/Content/Media//daily-shipping-and-cargo-handling-reports/"
              "dailyShippingCargoStat_10-OCT-2026-484958aebe2341e5a83f88097913ded5.PDF",):
        rr = requests.get(f, headers=H, timeout=(15, 120), verify=False)
        out(f"PQA daily PDF {rr.status_code} {len(rr.content)}b")
        pdf_text("PQA", rr.content, pat=r"LNG|EETL|PGPC", pages=4, chars=3500)
        import pdfplumber
        with pdfplumber.open(io.BytesIO(rr.content)) as pdf:
            for i, pg in enumerate(pdf.pages):
                for ln in (pg.extract_text() or "").splitlines():
                    if re.search(r"LNG|EETL|PGPC", ln):
                        out(f"    LNG line p{i + 1}: {ln[:200]}")
    # ---- SBP LNG import payments: last columns
    import pandas as pd
    rr = requests.get("https://www.sbp.org.pk/assets/document/Import_Payments_by_Commodities_and_Groups_Arch.xls",
                      headers=H, timeout=(15, 180), verify=False)
    df = pd.read_excel(io.BytesIO(rr.content), sheet_name="Groups (Jul-23 onwards)", header=None)
    hdr = df.iloc[3].tolist()
    row = df[df[0].astype(str).str.contains("Natural Ga", na=False)].iloc[0].tolist()
    out(f"\nSBP LNG row label={row[0]!r}; last 6: " + ", ".join(f"{h}: {v}" for h, v in list(zip(hdr, row))[-6:]))


MONTHS = r"\b(Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar|Apr|May|Jun)[a-z]*\b"


def ocr_pdf(body, pages=2, chars=2500):
    import pypdfium2 as pdfium
    import pytesseract
    doc = pdfium.PdfDocument(body)
    for i in range(min(pages, len(doc))):
        img = doc[i].render(scale=2.5).to_pil()
        tx = pytesseract.image_to_string(img)
        out(f"  --- OCR page {i + 1} ({len(tx)} chars):\n" + tx[:chars])


def round5():
    import pdfplumber
    # ---- OGRA State of the Regulated Petroleum Industry 2024-25: gas / LNG section tables
    r = cf_get("https://www.ogra.org.pk/download/14125?lang=en")
    if r is not None:
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for i, pg in enumerate(pdf.pages):
                tx = pg.extract_text() or ""
                nm = len(re.findall(MONTHS, tx))
                key = re.search(r"MMCF|BCF|MMBTU|cargo|sector|field|province|Table|Figure", tx, re.I)
                if 20 <= i <= 105 and key:
                    head = " / ".join(l for l in tx.splitlines()[:3])
                    out(f"PIR p{i + 1}: months={nm} chars={len(tx)} :: {head[:160]}")
                    if nm >= 6 or re.search(r"sector.?wise|field.?wise|company.?wise|cargo", tx, re.I):
                        out("      " + tx[:1400].replace("\n", "\n      "))
    # ---- OGRA RLNG provisional price determinations: text layer / OCR (cargo volumes?)
    for doc in ("16628", "13151"):
        r = cf_get(f"https://www.ogra.org.pk/download/{doc}?lang=en")
        if r is None:
            continue
        out(f"\nRLNG determination {doc}: {r.headers.get('content-disposition')} {len(r.content)}b")
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            n = len(pdf.pages)
            tx = "\n".join((p.extract_text() or "") for p in pdf.pages[:6])
        out(f"  pages={n} text chars={len(tx)}")
        if len(tx) > 200:
            out(tx[:4000])
        else:
            ocr_pdf(r.content, pages=min(n, 4), chars=3000)
    # ---- PQA daily statements: LNG berths (EETL / PGPCL) over the last ~2 months; oldest album page
    files = []
    for pgn in range(0, 7):
        u = "https://pqa.gov.pk/en/documents/album/daily-shipping-and-cargo-handling-reports" + (f"?page={pgn}" if pgn else "")
        try:
            rr = requests.get(u, headers=H, timeout=T, verify=False)
            files += [f for f in re.findall(r"""["']([^"']*dailyShippingCargoStat_[^"']+)["']""", rr.text) if f not in files]
        except Exception as e:  # noqa: BLE001
            out(f"PQA {u}: {e}")
    for pgn in (100, 200, 245, 246):
        rr = requests.get(f"https://pqa.gov.pk/en/documents/album/daily-shipping-and-cargo-handling-reports?page={pgn}",
                          headers=H, timeout=T, verify=False)
        f = re.findall(r"""["']([^"']*\.(?:PDF|pdf|xlsx?))["']""", rr.text)
        dates = re.findall(r"\d{1,2} [A-Z][a-z]{2} 20\d\d", text_of(rr.text, 100000))
        out(f"PQA album page {pgn}: {rr.status_code} files={len(f)} first={f[:1]} dates={dates[:3]}..{dates[-2:]}")
    out(f"PQA: {len(files)} daily files on pages 0-6")
    for f in files:
        u = urljoin("https://www.pqa.gov.pk/", f)
        try:
            rr = requests.get(u, headers=H, timeout=(15, 90), verify=False)
            with pdfplumber.open(io.BytesIO(rr.content)) as pdf:
                tx = (pdf.pages[0].extract_text() or "")
            lines = tx.splitlines()
            hit = [f"{ln[:150]} || {lines[k + 1][:60] if k + 1 < len(lines) else ''}" for k, ln in enumerate(lines)
                   if re.match(r"\s*(EETL|PGPCL)\b", ln) and "VACANT" not in ln]
            d = re.search(r"dailyShippingCargoStat_([0-9A-Z-]+?)-[0-9a-f]{20,}", f)
            if hit:
                out(f"  {d.group(1) if d else f}: " + " ## ".join(hit))
        except Exception as e:  # noqa: BLE001
            out(f"  {f}: {type(e).__name__}")


if __name__ == "__main__":
    {"1": round1, "2": round2, "3": round3, "4": round4, "5": round5}[sys.argv[1] if len(sys.argv) > 1 else "1"]()
