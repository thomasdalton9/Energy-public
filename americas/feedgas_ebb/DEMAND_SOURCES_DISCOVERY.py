"""
Discovery for two demand proxies to sit alongside lng_feedgas_daily.py:

1. Texas/Louisiana pipeline offtake - the same OA postings' delivery
   meters, but to end users / LDCs rather than LNG terminals. Needs
   each meter's state and type (end user, LDC, interconnect...), which
   the OA postings mostly lack. Checks each platform's location data:
     - Kinder Morgan: LocationDataDownload/LocDataDwnld.aspx?code=X
     - Enbridge LINK: a few guessed posting types (the OA page itself
       links none; Texas Eastern's zone codes STX/ETX/WLA/ELA already
       split Texas from Louisiana as a fallback)
     - Energy Transfer: the OA CSV's own columns (the page legend lists
       State, County, Operator) and the master delivery point list

2. Power burn - gas-fired generation, converted to Bcf/d:
     - ERCOT's public fuel-mix dashboard JSON (no key)
     - MISO's real-time fuel mix JSON and historical fuel mix workbook
       (looking for a South/Louisiana region split)

Prints structure (keys, columns, sample rows) for each so the real
pull can be written against it. Outputs: demand_sources_output/.
"""

import io
import json
import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demand_sources_output")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
YESTERDAY = datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)


def log(msg=""):
    print(msg, flush=True)


def describe_json(obj, depth=0, max_depth=4):
    pad = "  " * depth
    if depth > max_depth:
        return
    if isinstance(obj, dict):
        keys = list(obj)
        log(f"{pad}dict with {len(keys)} keys: {keys[:12]}")
        for k in keys[:3]:
            log(f"{pad}[{k!r}]:")
            describe_json(obj[k], depth + 1, max_depth)
    elif isinstance(obj, list):
        log(f"{pad}list of {len(obj)}")
        if obj:
            describe_json(obj[0], depth + 1, max_depth)
    else:
        log(f"{pad}{type(obj).__name__}: {str(obj)[:120]!r}")


def get(url, **kw):
    try:
        r = requests.get(url, headers=HEADERS, timeout=60, **kw)
        log(f"  GET {url} -> HTTP {r.status_code}, {len(r.content):,} bytes, {r.headers.get('Content-Type')}")
        return r
    except requests.RequestException as e:
        log(f"  GET {url} FAILED: {type(e).__name__}: {e}")
        return None


def power_sources():
    log("\n==================== ERCOT fuel mix ====================")
    r = get("https://www.ercot.com/api/1/services/read/dashboards/fuel-mix.json")
    if r is not None and r.ok:
        try:
            data = r.json()
            with open(os.path.join(OUTPUT_DIR, "ercot_fuel_mix.json"), "w") as f:
                json.dump(data, f)
            describe_json(data, max_depth=5)
        except ValueError:
            log(f"  not JSON: {r.text[:300]!r}")

    log("\n==================== MISO real-time fuel mix ====================")
    r = get("https://api.misoenergy.org/MISORTWDDataBroker/DataBrokerServices.asmx?messageType=getfuelmix&returnType=json")
    if r is not None and r.ok:
        try:
            describe_json(r.json(), max_depth=4)
        except ValueError:
            log(f"  not JSON: {r.text[:300]!r}")

    log("\n==================== MISO historical fuel mix ====================")
    for year in (YESTERDAY.year, YESTERDAY.year - 1):
        r = get(f"https://docs.misoenergy.org/marketreports/historical_gen_fuel_mix_{year}.xlsx")
        if r is not None and r.ok and r.content[:2] == b"PK":
            with open(os.path.join(OUTPUT_DIR, f"miso_historical_gen_fuel_mix_{year}.xlsx"), "wb") as f:
                f.write(r.content)
            book = pd.ExcelFile(io.BytesIO(r.content))
            log(f"  sheets: {book.sheet_names}")
            raw = pd.read_excel(book, book.sheet_names[0], header=None, nrows=12)
            log(raw.to_string())
            break
    for url in ("https://docs.misoenergy.org/marketreports/"
                f"{YESTERDAY:%Y%m%d}_sr_gfm.xlsx",
                f"https://docs.misoenergy.org/marketreports/{YESTERDAY:%Y%m%d}_df_al.xls"):
        r = get(url)


def km_locations(page, code):
    url = f"https://pipeline2.kindermorgan.com/LocationDataDownload/LocDataDwnld.aspx?code={code}"
    log(f"\n--- KM location data {code}: {url}")
    try:
        page.goto(url, wait_until="networkidle", timeout=60000)
    except Exception as e:
        log(f"  load failed: {e}")
        return
    page.screenshot(path=os.path.join(OUTPUT_DIR, f"km_locdata_{code}.png"), full_page=True)
    controls = page.evaluate("""() => [...document.querySelectorAll('input,select,button,a')]
        .filter(e => e.type !== 'hidden')
        .map(e => `${e.tagName} id=${e.id} type=${e.type || ''} text=${(e.innerText || e.value || e.alt || '').trim().slice(0, 50)}`)
        .filter(s => /download|retrieve|excel|csv|export|btn/i.test(s)).slice(0, 25)""")
    for c in controls:
        log(f"  {c}")
    log(f"  text: {page.inner_text('body')[:600]!r}")
    for sel in ("input[id$='btnDownload']", "input[id$='btnRetrieve']", "a:has-text('Download')", "input[value*='Download']"):
        if page.locator(sel).count():
            log(f"  trying {sel}")
            try:
                with page.expect_download(timeout=60000) as info:
                    page.locator(sel).first.click()
                path = os.path.join(OUTPUT_DIR, f"km_locdata_{code}_{info.value.suggested_filename}")
                info.value.save_as(path)
                log(f"  DOWNLOADED {path}")
                try:
                    df = pd.read_excel(path, header=None, nrows=8) if path.endswith(("xls", "xlsx")) else pd.read_csv(path, nrows=8, header=None)
                    log(df.to_string())
                except Exception as e:
                    log(f"  couldn't parse: {e}")
                return
            except Exception as e:
                log(f"  no download: {type(e).__name__}: {str(e)[:150]}")
                page.wait_for_load_state("networkidle")
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"km_locdata_{code}_after_click.png"), full_page=True)


def enbridge_locations(page):
    for t in ("LOC", "LD", "LDD", "LOCDATA", "Locations", "LDL"):
        url = f"https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu=TE&Type={t}"
        log(f"\n--- Enbridge {url}")
        try:
            page.goto(url, wait_until="networkidle", timeout=60000)
        except Exception as e:
            log(f"  load failed: {e}")
            continue
        text = page.inner_text("body")
        log(f"  text: {' '.join(text.split())[:300]!r}")
        links = page.evaluate("""() => [...document.querySelectorAll('a')].map(a => (a.innerText||'').trim()).filter(Boolean).slice(0, 30)""")
        log(f"  links: {links}")


def et_columns(page):
    for asset in ("TGC", "GR"):
        page_url = f"https://tgcmessenger.energytransfer.com/ipost/capacity/operationally-available-by-location?asset={asset}"
        csv_url = page_url + f"&f=csv&extension=csv&gasDay={YESTERDAY:%m}%2F{YESTERDAY:%d}%2F{YESTERDAY:%Y}&cycleDesc=Final&pointCd=&name="
        log(f"\n--- Energy Transfer {asset} OA CSV columns")
        page.goto(page_url, wait_until="networkidle", timeout=60000)
        r = page.context.request.get(csv_url, timeout=90000)
        df = pd.read_csv(io.StringIO(r.text()))
        log(f"  columns: {list(df.columns)}")
        log(df.head(5).to_string())
        for col in df.columns:
            if col.strip().lower() in ("state", "st"):
                log(f"  state counts: {df[col].value_counts().head(10).to_dict()}")
    log("\n--- Energy Transfer master delivery point list (TGC)")
    page.goto("https://tgcmessenger.energytransfer.com/ipost/main/index?asset=TGC", wait_until="networkidle", timeout=60000)
    links = page.evaluate("""() => [...document.querySelectorAll('a')].filter(a => /point list|point catalog|location/i.test(a.innerText))
        .map(a => a.innerText.trim() + ' -> ' + a.href)""")
    for l in links:
        log(f"  {l}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    power_sources()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()
        for code in ("NGPL", "TGP"):
            km_locations(page, code)
        enbridge_locations(page)
        et_columns(page)
        browser.close()
    log(f"\nDONE. Outputs in {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
