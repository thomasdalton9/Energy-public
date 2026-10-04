"""Probe 9 (probe 8 got the F5 challenge page; this one polls past it): transd.nea.org.np 'Transmission Operational Reports' lists Daily / Monthly / Yearly Operational
Report. Follow 'Daily Operational Report', print its listing (links, pagination) and the text of a few PDFs."""
import io
import re
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def strip(h):
    h = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>|<image[^>]*>", " ", h, flags=re.S)
    return re.sub(r"\s+", " ", h)


def anchors(html, base):
    return [(urljoin(base, h), re.sub(r"<[^>]+>|\s+", " ", t).strip())
            for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, flags=re.S)]


with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True, accept_downloads=True)

    def visit(url, wait=4000, tries=4, marker=None):
        """Load a page past the F5 challenge: poll until the real page (marker text, or no challenge script) shows."""
        for k in range(tries):
            page = ctx.new_page()
            try:
                r = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                html = ""
                for _ in range(20):
                    page.wait_for_timeout(1500)
                    html = page.content()
                    challenge = "bobcmn" in html or "apm_do_not_touch" in html or "TSPD/?type" in html[:3000]
                    if (marker and marker in html) or (not marker and not challenge and len(html) > 15000):
                        break
                else:
                    print(f"  (still challenge after 30 s, retry {k}) {url} len {len(html)}")
                    page.close()
                    continue
                page.wait_for_timeout(wait)
                html = page.content()
                print(f"\n=== {r.status if r else '?'} {page.url} len {len(html)} title {page.title()}", flush=True)
                return page, html
            except Exception as e:
                print("VISIT ERR", url, type(e).__name__, str(e)[:160])
                page.close()
        return None, ""

    def pdf(url):
        for k in range(3):
            try:
                r = ctx.request.get(url, timeout=90000)
                body = r.body()
                print(f"PDF GET {r.status} {len(body)} {r.headers.get('content-type')} {url}", flush=True)
                if body[:4] == b"%PDF":
                    return body
            except Exception as e:
                print("PDF ERR", type(e).__name__, str(e)[:150])
        return None

    page, html = visit("https://transd.nea.org.np/en/pages/transmission-operational-reports", marker="Daily Operational")
    s = strip(html)
    i = s.find("S.N.")
    print("TABLE HTML:", s[max(0, i - 200):i + 2500])
    links = [(u, t) for u, t in anchors(html, page.url if page else "")
             if re.search(r"operational|daily|monthly|yearly", t + u, re.I)]
    for u, t in links:
        print("  LINK", t, "->", u)
    if page:
        page.close()
    daily = [u for u, t in links if re.search(r"daily", t + u, re.I)]
    for du in daily[:2]:
        if du.lower().endswith(".pdf"):
            continue
        page, html = visit(du, 5000)
        s = strip(html)
        print("DAILY PAGE TEXT:", re.sub(r"<[^>]+>", " ", s)[:4000])
        al = anchors(html, page.url if page else du)
        for u, t in al:
            if re.search(r"pdf|download|uploads|storage|page=|detail|NDOR", u + t, re.I):
                print("   A", t[:80], "->", u)
        for x in re.findall(r'(?:src|data-src|data)="([^"]+\.pdf[^"]*)"', html):
            print("   EMBED", x)
        pdfs = [u for u, t in al if ".pdf" in u.lower()] + re.findall(r'(?:src|data)="([^"]+\.pdf[^"]*)"', html)
        for u in pdfs[:3]:
            body = pdf(urljoin(du, u))
            if body:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(body)) as d:
                    print("pages", len(d.pages))
                    for pg in d.pages[:2]:
                        print("-----"); print(pg.extract_text())
                        for t in pg.extract_tables()[:4]:
                            print("TABLE", t)
        if page:
            page.close()
    b.close()
