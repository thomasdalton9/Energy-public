"""
Third-pass discovery for Kinder Morgan's "Operationally Available
capacity by point" page (OpAvailPoint.aspx). The saved HTML from the
second pass revealed why a plain `requests.get()` never sees any data:
the grid (id="...DGOpAvail") is an Infragistics WebDataGrid that ships
as an EMPTY shell in the initial server response and only gets
populated client-side, in the browser, after clicking "Retrieve Data".
A real browser has to click the button and let its JavaScript run.

First GitHub run (28 Sep, ~5am Central): the page loaded and Retrieve
worked - the grid came back with exactly the columns wanted (Loc, Loc
Name, Design/Operating Capacity, Total Scheduled Quantity, Operationally
Available Capacity, ...) - but "No record(s) were found" for the
defaults (today's gas day, BEST AVAILABLE cycle, Delivery locations).
Download on that empty grid went to KM's generic error page.

So this pass tries every combination of:
  - gas day: yesterday, the day before, today (Central time)
  - location type: Delivery points, Receipt points
and records the row count / grid text / a screenshot for each. Then, for
the first combination that returned rows, it re-runs that query and
tries the "Download Data" button (EXCEL is the default format in its
dropdown) - an actual file export would be far easier to parse than
the rendered grid.

Needs Playwright + Chromium - the feedgas_ebb_discovery.yml workflow
installs both on GitHub's runner. Outputs go to
km_retrieve_discovery_output/ next to this script.
"""

import os
import re
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print(
        "Playwright isn't installed in this environment. Run:\n"
        "  pip install playwright\n"
        "  playwright install chromium\n"
        "then re-run this script.",
        file=sys.stderr,
    )
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "km_retrieve_discovery_output")
URL = "https://pipeline2.kindermorgan.com/Capacity/OpAvailPoint.aspx?code=GCX"

PREFIX = "#WebSplitter1_tmpl1_ContentPlaceHolder1_"
RETRIEVE_BTN = PREFIX + "HeaderBTN1_btnRetrieve"
DOWNLOAD_BTN = PREFIX + "HeaderBTN1_btnDownload"
GRID_CONTAINER = PREFIX + "DGOpAvail"
LOCATION_RADIOS = {"delivery": PREFIX + "rbDelivery", "receipt": PREFIX + "rbReceipt"}
# The "Eff Gas Day" Infragistics date editor has no stable id - it's the
# only input with this class on the page (seen in the first run's HTML).
GAS_DAY_INPUT = "input.igte_NautilusEditInContainer"


def gas_days():
    today = datetime.now(ZoneInfo("America/Chicago")).date()
    return [today - timedelta(days=1), today - timedelta(days=2), today]


def fmt(day):
    return f"{day.month}/{day.day}/{day.year}"  # the page's own M/D/YYYY format


def run_query(page, day, location):
    """Set the gas day and location type, click Retrieve, return grid text."""
    page.click(LOCATION_RADIOS[location])
    page.wait_for_load_state("networkidle", timeout=30000)  # the radio may post back

    box = page.locator(GAS_DAY_INPUT).first
    box.click()
    page.keyboard.press("Control+A")
    page.keyboard.type(fmt(day))
    page.keyboard.press("Tab")
    page.wait_for_timeout(500)

    page.click(RETRIEVE_BTN, timeout=10000)
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(2000)  # let the client-side grid render
    return page.inner_text(GRID_CONTAINER, timeout=5000)


def row_count(grid_text):
    m = re.search(r"Row Count:\s*(\d+)", grid_text)
    return int(m.group(1)) if m else None


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1400, "height": 1000})
        page = context.new_page()

        print(f"Navigating to {URL} ...", flush=True)
        page.goto(URL, wait_until="networkidle", timeout=30000)

        for day in gas_days():
            for location in LOCATION_RADIOS:
                tag = f"{day.isoformat()}_{location}"
                print(f"\n--- gas day {fmt(day)}, {location} points ---", flush=True)
                try:
                    text = run_query(page, day, location)
                except Exception as e:
                    print(f"  FAILED: {type(e).__name__}: {e}", flush=True)
                    page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}_FAILED.png"), full_page=True)
                    results.append((day, location, None))
                    continue
                shown_day = page.locator(GAS_DAY_INPUT).first.input_value()
                n = row_count(text)
                print(f"  gas day box now shows {shown_day!r}; row count: {n}", flush=True)
                print(f"  first 600 chars of grid text:\n{text[:600]}", flush=True)
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
                with open(os.path.join(OUTPUT_DIR, f"{tag}_grid.txt"), "w", encoding="utf-8") as f:
                    f.write(text)
                results.append((day, location, n))

        print("\n=== SUMMARY ===", flush=True)
        for day, location, n in results:
            print(f"  {fmt(day):>10}  {location:<8}  rows: {n}", flush=True)

        hit = next(((d, loc) for d, loc, n in results if n), None)
        if hit is None:
            print("\nNo combination returned rows - skipping the download attempt.", flush=True)
        else:
            day, location = hit
            print(f"\nRe-running {fmt(day)} {location} and trying Download (EXCEL)...", flush=True)
            run_query(page, day, location)
            with open(os.path.join(OUTPUT_DIR, "grid_with_rows.html"), "w", encoding="utf-8") as f:
                f.write(page.inner_html(GRID_CONTAINER))
            try:
                with page.expect_download(timeout=20000) as download_info:
                    page.click(DOWNLOAD_BTN, timeout=10000)
                download = download_info.value
                save_path = os.path.join(OUTPUT_DIR, "download_" + download.suggested_filename)
                download.save_as(save_path)
                print(f"  File downloaded -> {save_path} ({os.path.getsize(save_path):,} bytes)", flush=True)
            except Exception as e:
                print(f"  No download captured ({type(e).__name__}: {e})", flush=True)
                page.wait_for_timeout(1500)
                page.screenshot(path=os.path.join(OUTPUT_DIR, "download_attempt.png"), full_page=True)
                for i, pg in enumerate(context.pages):
                    if pg is not page:
                        pg.screenshot(path=os.path.join(OUTPUT_DIR, f"download_popup_{i}.png"), full_page=True)

        browser.close()

    print(f"\nDONE. Outputs in {OUTPUT_DIR}/", flush=True)


if __name__ == "__main__":
    main()
