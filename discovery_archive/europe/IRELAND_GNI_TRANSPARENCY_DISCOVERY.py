"""
Read GNI's data-transparency pages (which should carry daily entry flows
by entry point and gas consumption by market sector up to the present,
unlike the quarterly open-data CSVs that end 2026-03-31):
  https://www.gasnetworks.ie/about/data-transparency/entry-flows/physical-flows
  https://www.gasnetworks.ie/about/data-transparency/exit-flows/gas-consumption-by-market-sector
Pass 1: raw HTML - every href/src that looks like data (csv/xlsx/json/api/
download/iframe). Pass 2: Playwright - load each page, capture every
network response that is JSON/CSV/XLSX or whose URL mentions api/data/
export, print its URL, status, content-type and a body sample; also try
to click anything labelled download/export/csv/excel and report the
resulting request. Not reachable from the editing sandbox.
"""
import json
import re
import sys

import requests

PAGES = [
    "https://www.gasnetworks.ie/about/data-transparency/entry-flows/physical-flows",
    "https://www.gasnetworks.ie/about/data-transparency/exit-flows/gas-consumption-by-market-sector",
]
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
DATA_RE = re.compile(r"(csv|xlsx?|json|api|download|export|iframe|powerbi|tableau|highcharts|chart|data)", re.I)


def out(*a):
    print(*a, flush=True)


out("################ PASS 1: raw HTML ################")
for url in PAGES:
    r = requests.get(url, headers=HEADERS, timeout=(10, 60))
    out(f"\n=== {url} -> {r.status_code}, {len(r.text)} chars")
    html = r.text
    for m in sorted(set(re.findall(r'(?:href|src|data-src|data-url|action)=["\']([^"\']+)["\']', html))):
        if DATA_RE.search(m):
            out("  link:", m)
    for m in sorted(set(re.findall(r'https?://[^\s"\'<>]+', html))):
        if DATA_RE.search(m) and "gasnetworks.ie/about" not in m:
            out("  url :", m)
    for m in re.findall(r"<iframe[^>]+>", html, re.I):
        out("  iframe:", m[:300])
    scripts = re.findall(r"<script[^>]*src=[\"']([^\"']+)[\"']", html, re.I)
    out(f"  {len(scripts)} script srcs:", [s for s in scripts if "gasnetworks" in s or s.startswith("/")][:30])
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    i = text.lower().find("flow") if "physical" in url else text.lower().find("sector")
    out("  text sample:", text[max(0, i - 400): i + 1600])

out("\n################ PASS 2: Playwright network capture ################")
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch()
    for url in PAGES:
        ctx = browser.new_context(user_agent=HEADERS["User-Agent"], accept_downloads=True)
        page = ctx.new_page()
        seen = []

        def on_response(resp):
            try:
                ct = resp.headers.get("content-type", "")
                u = resp.url
                if any(k in ct for k in ("json", "csv", "excel", "spreadsheet", "octet")) or \
                        re.search(r"(api|data|export|download|flows|consumption|sector|chart)", u, re.I):
                    if u.endswith((".js", ".css", ".png", ".svg", ".woff", ".woff2", ".jpg", ".gif")):
                        return
                    body = ""
                    try:
                        body = resp.text()[:600]
                    except Exception:
                        pass
                    seen.append((u, resp.status, ct, body))
            except Exception:
                pass

        page.on("response", on_response)
        out(f"\n=== {url}")
        try:
            page.goto(url, timeout=90000, wait_until="load")
            page.wait_for_timeout(8000)
        except Exception as e:
            out("  goto error:", type(e).__name__, e)
        out("  title:", page.title())
        for u, st, ct, body in seen:
            out(f"  RESP {st} {ct[:40]} {u}")
            if body:
                out("       body:", body.replace("\n", " ")[:600])
        # buttons / links that might trigger an export
        cands = page.locator("a, button").filter(has_text=re.compile(r"download|export|csv|excel|xlsx", re.I))
        n = cands.count()
        out(f"  {n} download/export-looking controls")
        for i in range(min(n, 8)):
            el = cands.nth(i)
            try:
                out("   control:", (el.inner_text() or "")[:80].strip(), "| href=", el.get_attribute("href"))
                before = len(seen)
                with page.expect_download(timeout=8000) as dl:
                    el.click(timeout=5000)
                d = dl.value
                out("     -> download:", d.suggested_filename, d.url)
                path = d.path()
                with open(path, "rb") as f:
                    out("     head:", f.read(400))
            except Exception as e:
                out("     (no download)", type(e).__name__, str(e)[:120])
                for u, st, ct, body in seen[before:]:
                    out(f"     new RESP {st} {ct[:40]} {u}")
        # date pickers / selects that hint at history
        for sel in page.locator("select, input[type=date], input[type=text]").all()[:10]:
            try:
                out("   input:", sel.get_attribute("name"), sel.get_attribute("id"), sel.get_attribute("placeholder"))
            except Exception:
                pass
        ctx.close()
    browser.close()
