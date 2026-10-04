"""Probe 6: new nea.org.np - sitemap/robots, category list, home-page scripts feeding 'Energy Details', search
form, ptddms.nea.org.np (Power Trade DMS), and whether any old /admin/assets/uploads file still resolves."""
import re
from playwright.sync_api import sync_playwright
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)
    net = []

    def visit(url, wait=5000, tries=3):
        for k in range(tries):
            page = ctx.new_page()
            page.on("response", lambda r: net.append((r.status, r.request.resource_type, r.url)))
            try:
                r = page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(wait)
                html = page.content()
                if "bobcmn" in html or len(html) < 9000 and "TSPD" in html:
                    print(f"  (challenge page, retry {k}) {url}")
                    page.close()
                    continue
                print(f"\n=== {r.status if r else '?'} {page.url} len {len(html)} title {page.title()}", flush=True)
                return page, html
            except Exception as e:
                print("VISIT ERR", url, type(e).__name__, str(e)[:160])
                page.close()
        return None, ""

    page, home = visit("https://nea.org.np/en", 7000)
    for s in re.findall(r"<script[^>]*>(.*?)</script>", home, flags=re.S):
        if re.search(r"ajax|fetch\(|/api|energy|axios", s, re.I) and "bobcmn" not in s:
            print("SCRIPT:", re.sub(r"\s+", " ", s)[:1500])
    for m in re.finditer(r"<form[^>]*>.*?</form>", home, flags=re.S):
        print("FORM:", re.sub(r"\s+", " ", m.group(0))[:600])
    if page:
        page.close()
    for u in ("https://nea.org.np/robots.txt", "https://nea.org.np/sitemap.xml", "https://nea.org.np/en/category",
              "https://nea.org.np/category", "https://ptddms.nea.org.np/", "https://ldc.nea.org.np/reports",
              "https://ldc.nea.org.np/downloads", "https://nea.org.np/en/pages/system-operation-department",
              "https://nea.org.np/en/category/ldc", "https://nea.org.np/en/category/load-dispatch-center",
              "https://nea.org.np/en/category/daily-operational-report", "https://nea.org.np/en/category/reports"):
        page, html = visit(u, 3500)
        if not page:
            continue
        try:
            t = re.sub(r"\s+", " ", page.inner_text("body"))
        except Exception:
            t = html
        print("TEXT", t[:1800])
        for l in sorted(set(re.findall(r'(?:href|action)="([^"]+)"', html))):
            if re.search(r"category/|/pages/|pdf|report|ldc|loc>|operation|energy|api", l, re.I):
                print("   ->", l)
        for l in re.findall(r"<loc>([^<]+)</loc>", html)[:400]:
            print("   LOC", l)
        page.close()
    # old file store
    for u in ("https://www.nea.org.np/admin/assets/uploads/ldc/NMOR%202080_01.pdf",
              "https://nea.org.np/admin/assets/uploads/ldc/NMOR%202080_01.pdf"):
        r = ctx.request.get(u, timeout=60000)
        print("OLD", r.status, len(r.body()), r.headers.get("content-type"), u)
    print("\nNET (non-static):")
    for s, t, u in net:
        if t in ("xhr", "fetch") and "TSPD" not in u:
            print("  ", s, t, u[:200])
    b.close()
