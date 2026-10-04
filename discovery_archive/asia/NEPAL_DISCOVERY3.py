"""Probe 3: nea.org.np sits behind an F5 BIG-IP ASM bot challenge (TSPD, 5.6 KB 'bobcmn' page to plain requests).
Load it in headless Chromium (playwright), log the SPA's XHR/API calls, find the LDC report pages, then fetch NDOR
PDFs through the browser context (its cookies pass the challenge)."""
import io
import re
import json
from playwright.sync_api import sync_playwright

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
calls = []


def show_pdf(body, label):
    import pdfplumber
    print("PDF OK", label, len(body), flush=True)
    with pdfplumber.open(io.BytesIO(body)) as d:
        print("pages", len(d.pages))
        for pg in d.pages[:2]:
            print("-----"); print(pg.extract_text())
            for t in pg.extract_tables()[:3]:
                print("TABLE", json.dumps(t)[:3000])


with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True, accept_downloads=True)
    page = ctx.new_page()

    def onresp(r):
        u = r.url
        if r.request.resource_type in ("xhr", "fetch", "document") or "/api" in u:
            calls.append(u)
            ct = r.headers.get("content-type", "")
            print(f"  NET {r.status} {r.request.resource_type} {ct[:30]} {u[:200]}", flush=True)
            if "json" in ct:
                try:
                    print("     JSON", r.text()[:800].replace("\n", " "))
                except Exception:
                    pass
    page.on("response", onresp)
    for url in ("https://www.nea.org.np/", "https://www.nea.org.np/"):
        try:
            page.goto(url, wait_until="networkidle", timeout=90000)
        except Exception as e:
            print("goto", url, type(e).__name__, str(e)[:200])
        page.wait_for_timeout(5000)
        print("TITLE", page.title(), "len", len(page.content()), flush=True)
    print("COOKIES", [c["name"] for c in ctx.cookies()])
    html = page.content()
    hrefs = sorted(set(re.findall(r'href="([^"]+)"', html)))
    print("N links", len(hrefs))
    for h in hrefs:
        if re.search(r"ldc|load|dispatch|report|daily|operation|download|publication|ndor|system", h, re.I):
            print("  HREF", h)
    txt = page.inner_text("body")
    print("BODY TEXT", re.sub(r"\s+", " ", txt)[:2500])
    for m in re.finditer(r"[^\n]{0,80}(Load Dispatch|Daily|LDC|System Operation)[^\n]{0,80}", txt):
        print("  TXT", m.group(0))

    # direct PDFs through the context (shares the challenge cookies)
    for host in ("https://www.nea.org.np", "https://nea.org.np"):
        for name in ("NDOR 2081_05_30.pdf", "NDOR 2080_03_15.pdf", "NDOR 2082_06_15.pdf"):
            for sep in ("\\", "/"):
                u = f"{host}/" + sep.join(["admin", "assets", "uploads", "ldc", name.replace(" ", "%20")])
                try:
                    r = ctx.request.get(u, timeout=60000)
                    body = r.body()
                    print(f"REQ {r.status} {len(body)} {r.headers.get('content-type','')[:30]} {u}", flush=True)
                    if body[:4] == b"%PDF":
                        show_pdf(body, u)
                        raise StopIteration
                except StopIteration:
                    break
                except Exception as e:
                    print("REQ ERR", u, type(e).__name__, str(e)[:150])
    # via page navigation too
    try:
        r = page.goto("https://www.nea.org.np/admin/assets/uploads/ldc/NDOR%202081_05_30.pdf", timeout=60000)
        print("NAV", r.status if r else None, r.headers.get("content-type") if r else None)
    except Exception as e:
        print("NAV ERR", type(e).__name__, str(e)[:200])
    # candidate SPA routes
    for route in ("/ldc", "/reports", "/load-dispatch-center", "/system-operation", "/publications", "/downloads"):
        try:
            page.goto("https://www.nea.org.np" + route, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(2000)
            t = re.sub(r"\s+", " ", page.inner_text("body"))
            print("ROUTE", route, page.url, t[:500])
        except Exception as e:
            print("ROUTE ERR", route, type(e).__name__, str(e)[:150])
    b.close()
print("API-ish calls:")
for u in sorted(set(calls)):
    print("  ", u[:250])
