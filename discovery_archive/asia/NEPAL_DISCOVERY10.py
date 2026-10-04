"""Probe 10: the daily operational reports category ('daily-operational-reports-1', 64 listing pages on the vendor
host td.neasite.dryicesolutions.net, newest 2081-09-27). Check the same category on transd.nea.org.np and
nea.org.np, plain requests vs browser, a detail page's PDF link, PDF text, and the listing's oldest page."""
import io
import re
from urllib.parse import urljoin
import requests
import urllib3
from playwright.sync_api import sync_playwright
urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
CAT = "/en/category/daily-operational-reports-1"
HOSTS = ["https://transd.nea.org.np", "https://td.neasite.dryicesolutions.net", "https://nea.org.np"]


def details(html):
    return list(dict.fromkeys(re.findall(r'href="([^"]+/detail/[^"]*daily-operational[^"]*)"', html)))


def maxpage(html):
    p = [int(x) for x in re.findall(r"[?&]page=(\d+)", html)]
    return max(p) if p else None


for h in HOSTS:
    try:
        r = requests.get(h + CAT, headers={"User-Agent": UA}, timeout=60, verify=False)
        print("REQ", r.status_code, len(r.text), "challenge" if "bobcmn" in r.text or "apm_do_not_touch" in r.text
              else "", h + CAT, "details", details(r.text)[:3], "maxpage", maxpage(r.text), flush=True)
    except Exception as e:
        print("REQ ERR", h, type(e).__name__, str(e)[:150])


def pdftext(body):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(body)) as d:
        print("pages", len(d.pages))
        for pg in d.pages[:3]:
            print("-----"); print(pg.extract_text())
            for t in pg.extract_tables()[:6]:
                print("TABLE", t)


with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)

    def visit(url, marker=None, tries=3):
        for k in range(tries):
            page = ctx.new_page()
            try:
                r = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                html = ""
                for _ in range(20):
                    page.wait_for_timeout(1500)
                    html = page.content()
                    ch = "bobcmn" in html or "apm_do_not_touch" in html
                    if (marker and marker in html) or (not marker and not ch and len(html) > 15000):
                        break
                else:
                    print(f"  (challenge/marker missing, retry {k}) {url} len {len(html)}")
                    page.close()
                    continue
                print(f"=== {r.status if r else '?'} {page.url} len {len(html)}", flush=True)
                page.close()
                return html
            except Exception as e:
                print("VISIT ERR", url, type(e).__name__, str(e)[:160])
                page.close()
        return ""

    for h in HOSTS[:2]:
        html = visit(h + CAT, marker="/detail/")
        d = details(html)
        print("  DETAILS", len(d), d[:12], "maxpage", maxpage(html))
        if not d:
            continue
        last = maxpage(html)
        if last:
            html2 = visit(f"{h}{CAT}?page={last}", marker="/detail/")
            print("  OLDEST PAGE", details(html2))
        dhtml = visit(urljoin(h, d[0]), marker="</main>")
        s = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>", " ", dhtml, flags=re.S)
        i = s.find("<main")
        print("  DETAIL MAIN:", re.sub(r"\s+", " ", s[i:i + 3000]))
        files = list(dict.fromkeys(re.findall(r'(?:href|src|data)="([^"]+\.(?:pdf|xlsx?|PDF)[^"]*)"', dhtml)))
        print("  FILES", files)
        for f in files[:1]:
            fu = urljoin(h, f)
            rr = requests.get(fu, headers={"User-Agent": UA}, timeout=90, verify=False)
            print("  PLAIN GET", rr.status_code, len(rr.content), rr.headers.get("content-type"), fu)
            body = rr.content if rr.content[:4] == b"%PDF" else None
            if body is None:
                r2 = ctx.request.get(fu, timeout=90000)
                print("  CTX GET", r2.status, len(r2.body()), r2.headers.get("content-type"))
                body = r2.body() if r2.body()[:4] == b"%PDF" else None
            if body:
                pdftext(body)
    b.close()
