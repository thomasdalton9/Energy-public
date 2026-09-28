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
    # Sabine Pass Trains 1-2 commissioned in 2015, before the wording above
    "sabine2015": ("Sabine Pass Liquefaction commissioning", "2014-06-01", "2016-06-30"),
}


def log(msg=""):
    print(msg, flush=True)


API = "https://elibrary.ferc.gov/eLibraryWebAPI/api/Search/AdvancedSearch"


def api_body(text, start="2010-01-01", end="2026-12-31"):
    return {"searchText": text, "searchFullText": True, "searchDescription": True,
            "dateSearches": [{"dateType": "filed_date", "startDate": start, "endDate": end}],
            "availability": None, "affiliations": [], "categories": [], "libraries": [], "accessionNumber": None,
            "eFiling": False, "docketSearches": [], "resultsPerPage": 100, "curPage": 0, "classTypes": [],
            "sortBy": "", "groupBy": "NONE", "idolResultID": "", "allDates": False}


def print_hits(data):
    hits = data if isinstance(data, list) else (data.get("searchHits") or data.get("SearchHits") or
                                                data.get("results") or data.get("Results") or [])
    log(f"    {len(hits)} hits; keys: {list(data)[:15] if isinstance(data, dict) else 'list'}")
    for h in hits[:80]:
        flat = json.dumps(h)
        date = re.search(r'"(?:filedDate|issuedDate|docDate|FiledDate|IssuedDate|filed_date|issued_date)"\s*:\s*"([^"]+)"', flat, re.I)
        desc = re.search(r'"(?:description|Description|title|Title)"\s*:\s*"([^"]{0,300})', flat)
        dockets = re.findall(r"[CR]P\d\d-\d+", flat)
        log(f"    HIT {date.group(1)[:10] if date else '?'} {sorted(set(dockets))[:3]} {desc.group(1) if desc else flat[:250]}")


def run_query(page, tag, query):
    text, start, end = query if isinstance(query, tuple) else (query, "2010-01-01", "2026-12-31")
    log(f"\n==================== {tag}: {text!r} ({start} to {end}) ====================")
    # 1. the search API directly, every page; keep FERC's own issuances
    #    (staff letters) that grant hazardous-fluid / commissioning / service
    wanted = re.compile(r"hazardous fluid|commission|fuel gas|feed gas|in-service|in service|commence service|"
                        r"place into service|placed into service|start-?up", re.I)
    kept, total = [], None
    try:
        for cur in range(12):
            body = api_body(text, start, end)
            body["curPage"] = cur
            r = page.context.request.post(API, data=body, timeout=120000)
            if not r.ok:
                log(f"  API {r.status}: {r.text()[:300]!r}")
                break
            data = r.json()
            hits = data.get("searchHits") or []
            total = data.get("totalHits")
            for h in hits:
                if h.get("category") == "Issuance" and wanted.search(h.get("description") or ""):
                    kept.append(h)
            if len(hits) < 100:
                break
        kept.sort(key=lambda h: h["filedDate"][6:] + h["filedDate"][:5])
        log(f"  API: {total} results, {len(kept)} FERC issuances about hazardous fluids / commissioning / service:")
        for h in kept:
            log(f"    {h['filedDate']} {h.get('acesssionNumber')} {sorted({d[:9] for d in h.get('docketNumbers') or []})[:3]} "
                f"{h['description'][:400]}")
        with open(os.path.join(OUTPUT_DIR, f"{tag}_issuances.json"), "w") as f:
            json.dump(kept, f, indent=1)
        return
    except Exception as e:
        log(f"  API failed: {type(e).__name__}: {str(e)[:200]}")
    # 2. the search form, capturing the call it makes
    calls = []

    def on_response(resp):
        if "eLibraryWebAPI" in resp.url and resp.request.method == "POST":
            try:
                calls.append((resp.status, resp.url, resp.request.post_data, resp.text()))
            except Exception:
                pass
    page.on("response", on_response)
    try:
        page.goto(SEARCH_URL, wait_until="networkidle", timeout=120000)
        page.wait_for_timeout(3000)
        inputs = page.evaluate("""() => [...document.querySelectorAll('input, textarea, select, mat-select, button')].map(e => ({
            tag: e.tagName, type: e.type || '', id: e.id, name: e.name || '', ph: e.placeholder || '',
            aria: e.getAttribute('aria-label') || '', fc: e.getAttribute('formcontrolname') || '',
            text: (e.innerText || '').trim().slice(0, 40)}))""")
        for i in inputs:
            log(f"  CONTROL {i}")
        kw = next((i for i in inputs if re.search(r"keyword|searchtext|search text", " ".join([i["fc"], i["aria"], i["ph"], i["id"], i["name"]]), re.I)), None)
        if kw:
            sel = f"#{kw['id']}" if kw["id"] else f"[formcontrolname={kw['fc']}]"
            page.fill(sel, text)
        for i in inputs:
            key = " ".join([i["fc"], i["aria"], i["ph"], i["id"]]).lower()
            if "from" in key and ("date" in key or i["type"] == "text"):
                page.fill(f"#{i['id']}" if i["id"] else f"[formcontrolname={i['fc']}]", "01/01/2010")
        page.get_by_role("button", name=re.compile("^\\s*Search\\s*$", re.I)).first.click(timeout=15000)
        page.wait_for_load_state("networkidle", timeout=120000)
        page.wait_for_timeout(8000)
        page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
    except Exception as e:
        log(f"  form FAILED: {type(e).__name__}: {str(e)[:300]}")
    page.remove_listener("response", on_response)
    for status, url, post, body in calls:
        log(f"  CALL {status} POST {url}")
        log(f"    BODY {post[:2000] if post else ''}")
        try:
            print_hits(json.loads(body))
        except Exception:
            log(f"    {body[:300]!r}")


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
