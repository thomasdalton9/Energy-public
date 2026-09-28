"""
Scan Kinder Morgan's "Operationally Available capacity by point" page
(Capacity/OpAvailPoint.aspx?code=X) across many pipeline codes, looking
for ones that actually publish point-level rows - Total Scheduled
Quantity at an LNG terminal delivery point is effectively feedgas.

Why: KM_GCX_RETRIEVE_DISCOVERY.py got the query working on Gulf Coast
Express but every gas day / location type came back "No record(s)".
GCX is an intrastate Texas pipeline, outside FERC's posting rules, so
its page is likely always empty. Interstate pipes on the same platform
(NGPL, Tennessee Gas, Kinder Morgan Louisiana, ...) have to post.

For each code: load the page (unknown codes just error), then for
yesterday's gas day (Central time), BEST AVAILABLE cycle, and both
Delivery and Receipt location types: Retrieve, record the row count,
and save the grid text. Where rows come back, also try the Download
(EXCEL) button - a file export avoids the grid only rendering the rows
currently scrolled into view.

Usage: python3 KM_OPAVAIL_SCAN.py [CODE ...]   (default: CODES below)
Needs Playwright + Chromium; the feedgas_ebb_discovery.yml workflow
installs both. Outputs go to km_opavail_scan_output/ next to this script.
"""

import os
import re
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "km_opavail_scan_output")
URL = "https://pipeline2.kindermorgan.com/Capacity/OpAvailPoint.aspx?code={code}"

# Candidate KM platform codes - guesses beyond the three already seen
# working (GCX, NGPL, EEC) are cheap to try: a bad one just errors.
CODES = [
    "NGPL",  # Natural Gas Pipeline Co of America - Gulf Coast LNG deliveries
    "TGP",  # Tennessee Gas Pipeline - Sabine Pass / Cameron area
    "KMLP",  # Kinder Morgan Louisiana Pipeline - built to serve Sabine Pass
    "EEC",  # Elba Express - Elba Island LNG
    "SNG",  # Southern Natural Gas - Elba Island LNG
    "EPNG",  # El Paso Natural Gas
    "CIG",  # Colorado Interstate Gas
    "MEP",  # Midcontinent Express
    "FEP",  # Fayetteville Express
    "SLNG",  # Southern LNG - Elba Island terminal
    "GCX",  # control: known to return no rows
]
# First scan (27 Sep gas day): KMLP, EEC and MEP returned rows and Excel
# exports - KMLP's "SPLIQ/KMLP SP LIQUEFACTION CAMERON" point had 1.28
# million Dth scheduled into Sabine Pass. NGPL and the rest were caught
# mid-load, hence the wait in query().

PREFIX = "#WebSplitter1_tmpl1_ContentPlaceHolder1_"
RETRIEVE_BTN = PREFIX + "HeaderBTN1_btnRetrieve"
DOWNLOAD_BTN = PREFIX + "HeaderBTN1_btnDownload"
GRID_CONTAINER = PREFIX + "DGOpAvail"
TSP_NAME = PREFIX + "lblTSPNameValue"
LOCATION_RADIOS = {"delivery": PREFIX + "rbDelivery", "receipt": PREFIX + "rbReceipt"}
GAS_DAY_PICKER_ID = "WebSplitter1_tmpl1_ContentPlaceHolder1_dtePickerBegin"
GAS_DAY_INPUT = f"#{GAS_DAY_PICKER_ID} input.igte_NautilusEditInContainer"

# Typing into the Infragistics WebDatePicker garbles dates - set it
# through the control's own client API instead.
SET_DATE_JS = """([id, y, m, d]) => {
    const picker = $find(id);
    if (!picker) return null;
    picker.set_value(new Date(y, m - 1, d));
    return String(picker.get_value());
}"""


def fmt(day):
    return f"{day.month}/{day.day}/{day.year}"


def row_count(grid_text):
    m = re.search(r"Row Count:\s*(\d+)", grid_text)
    return int(m.group(1)) if m else None


def load(page, code):
    """Load the code's page; return its TSP name, or None if it isn't a
    working capacity page."""
    try:
        page.goto(URL.format(code=code), wait_until="networkidle", timeout=45000)
    except Exception as e:
        print(f"  load failed: {type(e).__name__}: {e}", flush=True)
        return None
    if page.locator(RETRIEVE_BTN).count() == 0:
        heading = page.locator("h1, h2").first.inner_text() if page.locator("h1, h2").count() else ""
        print(f"  no capacity form on the page ({heading.strip()[:80]!r})", flush=True)
        return None
    return page.locator(TSP_NAME).inner_text().strip() if page.locator(TSP_NAME).count() else "?"


def query(page, code, day, location):
    load(page, code)
    page.click(LOCATION_RADIOS[location])
    page.wait_for_load_state("networkidle", timeout=30000)
    if page.evaluate(SET_DATE_JS, [GAS_DAY_PICKER_ID, day.year, day.month, day.day]) is None:
        raise RuntimeError("$find() couldn't locate the gas day picker")
    page.wait_for_timeout(500)
    page.click(RETRIEVE_BTN, timeout=15000)
    page.wait_for_load_state("networkidle", timeout=45000)
    # Big systems (NGPL, TGP, ...) keep a "Loading... Please wait" overlay
    # long after the network goes quiet - the first scan screenshotted
    # NGPL mid-load and read an empty grid. Wait for the grid footer's
    # row count (or the no-records message) to actually appear.
    page.wait_for_function(
        """sel => { const g = document.querySelector(sel);
                    return g && /Row Count:\\s*\\d|No record/.test(g.innerText); }""",
        arg=GRID_CONTAINER,
        timeout=180000,
    )
    page.wait_for_timeout(1000)
    return page.inner_text(GRID_CONTAINER, timeout=10000)


def try_download(page, tag):
    try:
        with page.expect_download(timeout=30000) as info:
            page.click(DOWNLOAD_BTN, timeout=15000)
        download = info.value
        path = os.path.join(OUTPUT_DIR, f"{tag}_download_{download.suggested_filename}")
        download.save_as(path)
        print(f"    DOWNLOADED {path} ({os.path.getsize(path):,} bytes)", flush=True)
        return path
    except Exception as e:
        print(f"    no download ({type(e).__name__}: {str(e)[:150]})", flush=True)
        page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}_download_attempt.png"), full_page=True)
        return None


def main():
    codes = sys.argv[1:] or CODES
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    day = datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)
    summary = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()

        for code in codes:
            print(f"\n===== {code} =====", flush=True)
            tsp = load(page, code)
            if tsp is None:
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"{code}_no_page.png"))
                summary.append((code, "-", "no page", None, None))
                continue
            print(f"  TSP name: {tsp!r}", flush=True)

            for location in LOCATION_RADIOS:
                tag = f"{code}_{day.isoformat()}_{location}"
                try:
                    text = query(page, code, day, location)
                    shown = page.locator(GAS_DAY_INPUT).first.input_value()
                except Exception as e:
                    print(f"  {location}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
                    page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}_FAILED.png"), full_page=True)
                    summary.append((code, tsp, location, None, None))
                    continue
                n = row_count(text)
                print(f"  {location}: gas day box {shown!r}, rows {n}", flush=True)
                with open(os.path.join(OUTPUT_DIR, f"{tag}_grid.txt"), "w", encoding="utf-8") as f:
                    f.write(text)
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
                downloaded = None
                if n:
                    print(f"  first 1500 chars of grid text:\n{text[:1500]}", flush=True)
                    downloaded = try_download(page, tag)
                summary.append((code, tsp, location, n, downloaded))

        browser.close()

    print(f"\n=== SUMMARY (gas day {fmt(day)}, BEST AVAILABLE) ===", flush=True)
    for code, tsp, location, n, downloaded in summary:
        print(f"  {code:<7} {tsp[:40]:<40} {location:<9} rows: {n!s:<6} download: {os.path.basename(downloaded) if downloaded else '-'}", flush=True)


if __name__ == "__main__":
    main()
