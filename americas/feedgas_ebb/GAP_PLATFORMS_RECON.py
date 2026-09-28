"""
Recon for the feedgas gap platforms, with network capture - these are
single-page apps whose data comes from JSON calls the page makes, so
the most robust scraper is usually to call those endpoints directly.

  - Cheniere LNG Connection (lngconnection.cheniere.com): its home page
    dashboard already showed Creole Trail's "CT200111-CREOLE TRAIL-
    SPLIQ-D" delivery into Sabine Pass Liquefaction, but not which gas
    day. Here: open Capacity > Operationally Available for Creole Trail
    (company 200) and Corpus Christi (company 400), dump the tables and
    controls, and log every JSON/XHR response URL.
  - TC Energy (ANR, Columbia Gulf): the first attempt got connection
    resets from ebb.anrpl.com / ebb.tceconnects.com - retried here with
    the TCeConnects URL form found on the web, plus plain requests.
  - BHE GT&S (Cove Point): first attempt was cut off by a redirect from
    the page before; retried on its own.

Operator-hosted postings only (see EBB_PLATFORM_RECON.py). Outputs:
gap_platforms_recon_output/.
"""

import json
import os
import re
import sys

import requests

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gap_platforms_recon_output")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

TABLES_JS = """() => [...document.querySelectorAll('table')].map(t =>
    [...t.rows].slice(0, 25).map(r => [...r.cells].map(c => c.innerText.replace(/\\s+/g, ' ').trim())))
    .filter(t => t.length > 1)"""
CONTROLS_JS = """() => [...document.querySelectorAll('input,select,button')].filter(e => e.type !== 'hidden')
    .map(e => `${e.tagName} id=${e.id} name=${e.name || ''} type=${e.type || ''} value=${(e.value || '').slice(0, 30)}`
              + (e.tagName === 'SELECT' ? ' options=' + [...e.options].map(o => o.text).slice(0, 10).join('|') : ''))
    .slice(0, 40)"""


def log(msg=""):
    print(msg, flush=True)


def dump_page(page, tag):
    page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
    with open(os.path.join(OUTPUT_DIR, f"{tag}.html"), "w", encoding="utf-8") as f:
        f.write(page.content())
    log(f"  url: {page.url}")
    log(f"  text: {' '.join(page.inner_text('body').split())[:500]!r}")
    for c in page.evaluate(CONTROLS_JS):
        log(f"  CONTROL {c}")
    for t in page.evaluate(TABLES_JS):
        log(f"  TABLE ({len(t)} rows shown)")
        for r in t[:14]:
            log(f"    {r}")


def cheniere(context):
    page = context.new_page()
    api_calls = []

    def on_response(resp):
        ct = resp.headers.get("content-type", "")
        if "json" in ct or "/api/" in resp.url.lower() or resp.request.resource_type in ("xhr", "fetch"):
            try:
                body = resp.text()[:400]
            except Exception:
                body = "?"
            api_calls.append((resp.status, resp.request.method, resp.url, ct, body))

    page.on("response", on_response)
    log("\n==================== Cheniere LNG Connection ====================")
    page.goto("https://lngconnection.cheniere.com/", wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(4000)
    if page.locator("#cookieAcceptBtn").count():
        page.locator("#cookieAcceptBtn").click()
    for company, label in (("200", "creole_trail"), ("400", "corpus_christi")):
        log(f"\n--- {label} (companyradio={company})")
        try:
            page.locator(f"input[name='companyradio'][value='{company}']").check(force=True)
            page.wait_for_timeout(3000)
            dump_page(page, f"cheniere_{label}_home")
            page.locator("a", has_text="Operationally Available").first.click(force=True)
            page.wait_for_load_state("networkidle", timeout=60000)
            page.wait_for_timeout(5000)
            dump_page(page, f"cheniere_{label}_oac")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    log(f"\n--- Cheniere XHR/JSON responses ({len(api_calls)})")
    seen = set()
    for status, method, url, ct, body in api_calls:
        key = (method, re.sub(r"\d{4,}", "N", url))
        if key in seen:
            continue
        seen.add(key)
        log(f"  {status} {method} {url} [{ct}]")
        log(f"      {body[:300]!r}")
    with open(os.path.join(OUTPUT_DIR, "cheniere_api_calls.json"), "w") as f:
        json.dump(api_calls, f, indent=1)
    page.close()


def plain(url):
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        log(f"  requests GET {url} -> {r.status_code}, {len(r.content):,} bytes, title "
            f"{(re.search(r'<title>(.*?)</title>', r.text, re.S | re.I) or [None, ''])[1].strip()[:80]!r}")
    except requests.RequestException as e:
        log(f"  requests GET {url} FAILED: {type(e).__name__}: {str(e)[:150]}")


def simple(context, name, urls):
    log(f"\n==================== {name} ====================")
    for url in urls:
        plain(url)
        page = context.new_page()
        log(f"\n--- browser {url}")
        try:
            page.goto(url, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(4000)
            dump_page(page, f"{name}_{re.sub(r'[^A-Za-z0-9]+', '_', url)[8:80]}")
            links = page.evaluate("""() => [...document.querySelectorAll('a[href]')]
                .filter(a => /capacit|operational|avail|location/i.test(a.innerText + a.href))
                .map(a => a.innerText.trim().slice(0, 60) + ' -> ' + a.href).slice(0, 30)""")
            for l in links:
                log(f"  LINK {l}")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
        page.close()


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, user_agent=UA, viewport={"width": 1600, "height": 1200})
        cheniere(context)
        simple(context, "tcenergy", [
            "https://ebb.tceconnects.com/infopost/TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3037",
            "https://www.tceconnects.com/",
            "https://ebb.anrpl.com/",
        ])
        simple(context, "bhe", [
            "https://infopost.bhegts.com/",
            "https://www.bhegts.com/",
        ])
        browser.close()
    log(f"\nDONE. Outputs in {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
