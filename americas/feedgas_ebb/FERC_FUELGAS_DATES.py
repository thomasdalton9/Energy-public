"""
Find when FERC authorised each LNG export terminal (and each later phase)
to introduce fuel gas / hazardous fluids - the start-up dates the feedgas
workbook uses. FERC's staff letters granting "authorization to introduce
hazardous fluids" (or "commence commissioning") are public in eLibrary.

Drives eLibrary's own search page (elibrary.ferc.gov) in Chromium, one
query per terminal, and logs every result (date, docket, description) -
plus the search API call the page makes, so later runs can call it
directly.

Usage: python3 FERC_FUELGAS_DATES.py [terminal ...]   (default: all)
Outputs: ferc_fuelgas_output/ next to this script.
"""

import json
import os
import re
import sys

from playwright.sync_api import sync_playwright

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ferc_fuelgas_output")
SEARCH_URL = "https://elibrary.ferc.gov/eLibrary/search"

QUERIES = {
    "sabine": "Sabine Pass Liquefaction introduce hazardous fluids",
    "covepoint": "Cove Point liquefaction introduce hazardous fluids",
    "corpus": "Corpus Christi Liquefaction introduce hazardous fluids",
    "cameron": "Cameron LNG liquefaction introduce hazardous fluids",
    "freeport": "Freeport LNG liquefaction introduce hazardous fluids",
    "elba": "Elba Liquefaction introduce hazardous fluids",
    "calcasieu": "Calcasieu Pass introduce hazardous fluids",
    "plaquemines": "Plaquemines LNG introduce hazardous fluids",
    "goldenpass": "Golden Pass LNG introduce hazardous fluids",
}


def log(msg=""):
    print(msg, flush=True)


def run_query(page, tag, text):
    log(f"\n==================== {tag}: {text!r} ====================")
    calls = []

    def on_response(resp):
        if "eLibraryWebAPI" in resp.url or "api" in resp.url.lower() and resp.request.method == "POST":
            try:
                calls.append((resp.status, resp.request.method, resp.url, resp.request.post_data, resp.text()))
            except Exception:
                pass
    page.on("response", on_response)
    try:
        page.goto(SEARCH_URL, wait_until="networkidle", timeout=120000)
        page.wait_for_timeout(3000)
        box = page.locator("input[type=text], input[type=search], textarea").first
        box.fill(text)
        box.press("Enter")
        page.wait_for_load_state("networkidle", timeout=120000)
        page.wait_for_timeout(8000)
        page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
        body = " ".join(page.evaluate("() => document.body.innerText").split())
        log(f"  page text: {body[:1500]!r}")
    except Exception as e:
        log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    page.remove_listener("response", on_response)
    for status, method, url, post, text_ in calls:
        log(f"  CALL {status} {method} {url}")
        if post:
            log(f"    POST {post[:1500]}")
        log(f"    {text_[:300]!r}")
        try:
            data = json.loads(text_)
        except Exception:
            continue
        with open(os.path.join(OUTPUT_DIR, f"{tag}_results.json"), "w") as f:
            json.dump(data, f, indent=1)
        hits = data.get("searchHits") or data.get("SearchHits") or data.get("results") or []
        for h in hits[:60]:
            flat = json.dumps(h)
            date = re.search(r'"(?:filedDate|issuedDate|docDate|FiledDate|IssuedDate)"\s*:\s*"([^"]+)"', flat)
            desc = re.search(r'"(?:description|Description|title)"\s*:\s*"([^"]{0,400})', flat)
            dockets = re.findall(r"[CR]P\d\d-\d+", flat)
            log(f"    HIT {date.group(1) if date else '?'} {sorted(set(dockets))[:4]} {desc.group(1) if desc else flat[:300]}")


def main():
    wanted = sys.argv[1:] or list(QUERIES)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1600, "height": 1200},
                                      user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                                 "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
        page = context.new_page()
        for tag in wanted:
            run_query(page, tag, QUERIES[tag])
        browser.close()
    log(f"\nDONE. Outputs in {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
