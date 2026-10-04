"""Probe 7: the NDOR archive is gone from the new NEA site. Look at (a) the home-page 'Energy Details' widget in
full (which day does it show? any link/date?) in EN and NP, (b) genrd/transd 'operational reports' pages with
retries past the F5 challenge, (c) category/other-reports and publication-and-reports listings."""
import re
from playwright.sync_api import sync_playwright
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def strip(h):
    h = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>|<image[^>]*>", " ", h, flags=re.S)
    return re.sub(r"\s+", " ", h)


with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)

    def visit(url, wait=5000, tries=4):
        for k in range(tries):
            page = ctx.new_page()
            try:
                r = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(wait)
                html = page.content()
                if "bobcmn" in html or (len(html) < 9000 and "TSPD" in html):
                    print(f"  (challenge page, retry {k}) {url}")
                    page.close()
                    page = ctx.new_page()
                    continue
                print(f"\n=== {r.status if r else '?'} {page.url} len {len(html)} title {page.title()}", flush=True)
                return page, html
            except Exception as e:
                print("VISIT ERR", url, type(e).__name__, str(e)[:160])
                page.close()
        return None, ""

    for u in ("https://nea.org.np/en", "https://nea.org.np/np", "https://nea.org.np/"):
        page, html = visit(u, 6000)
        s = strip(html)
        i = s.find("Energy Details")
        if i < 0:
            i = s.find("ऊर्जा")
        print("WIDGET:", s[max(0, i - 300):i + 2500])
        if page:
            page.close()
    for u in ("https://genrd.nea.org.np/en/pages/generation-operational-reports",
              "https://transd.nea.org.np/en/pages/transmission-operational-reports",
              "https://nea.org.np/en/category/other-reports", "https://nea.org.np/en/category/publication-and-reports",
              "https://nea.org.np/en/category/auto-publication"):
        page, html = visit(u, 5000)
        s = strip(html)
        i = s.find("<main")
        print("MAIN:", s[i if i >= 0 else 0:][:3500])
        for l in sorted(set(re.findall(r'(?:href|src|data)="([^"]+)"', html))):
            if re.search(r"pdf|/detail/|storage|uploads|drive|iframe|page=", l, re.I):
                print("   ->", l)
        if page:
            for f in page.frames:
                print("   FRAME", f.url)
            page.close()
    b.close()
