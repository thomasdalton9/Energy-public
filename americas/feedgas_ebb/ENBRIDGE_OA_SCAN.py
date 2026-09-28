"""
Enbridge LINK "Operationally Available Capacity" CSV pull, per business
unit - the Enbridge counterpart of KM_OPAVAIL_SCAN.py.

EBB_PLATFORM_RECON.py showed the page structure at
rtba.enbridge.com/InformationalPosting/Default.aspx?bu=XX&Type=OA:
a Telerik date picker defaulting to yesterday's gas day, a cycle
dropdown (ddlSelector) defaulting to the latest posted cycle (e.g.
INTRDY_2026-09-28_0501), and a "Downloadable Format" link that posts
back and returns the postings as CSV. So: load, click the link, keep
the file - no form filling needed for the default (yesterday, latest
cycle).

Business units that touch LNG terminals:
  - TE  Texas Eastern - Freeport (2 of 3 meters), Calcasieu Pass
        (via East Lateral), Plaquemines (via Gator Express)
  - BIG BIG Pipeline - Freeport
  - AGT Algonquin - no export terminal, but carries Everett/Northeast
        LNG sendout; cheap to include

Usage: python3 ENBRIDGE_OA_SCAN.py [BU ...]   (default: BUS below)
Outputs: enbridge_oa_scan_output/ next to this script, plus a printout
of any row whose text mentions an LNG-ish name.
"""

import csv
import io
import os
import re
import sys

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "enbridge_oa_scan_output")
URL = "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu={bu}&Type=OA"
BUS = ["TE", "BIG", "AGT"]
DOWNLOAD_LINK = "a:has-text('Downloadable Format')"
CYCLE_SELECT = "#ddlSelector"
LNG_PATTERN = re.compile(r"LNG|LIQ|FREEPORT|CALCAS|PLAQUEM|GATOR|VENTURE|\bVG\b|CAMERON|SABINE|CORPUS|GOLDEN|STRATTON|COVE|ELBA", re.I)


def main():
    bus = sys.argv[1:] or BUS
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()
        for bu in bus:
            print(f"\n===== {bu} =====", flush=True)
            try:
                page.goto(URL.format(bu=bu), wait_until="networkidle", timeout=60000)
            except Exception as e:
                print(f"  load failed: {type(e).__name__}: {e}", flush=True)
                continue
            page.screenshot(path=os.path.join(OUTPUT_DIR, f"{bu}_page.png"), full_page=True)
            cycle = page.locator(CYCLE_SELECT).input_value() if page.locator(CYCLE_SELECT).count() else "?"
            gas_day = page.locator("#ctl00_MainContent_ctl01_oaDefault_ucDate_rdpDate_dateInput").input_value() \
                if page.locator("#ctl00_MainContent_ctl01_oaDefault_ucDate_rdpDate_dateInput").count() else "?"
            print(f"  gas day {gas_day!r}, cycle {cycle!r}", flush=True)
            if page.locator(DOWNLOAD_LINK).count() == 0:
                print("  no 'Downloadable Format' link on this page", flush=True)
                continue
            try:
                with page.expect_download(timeout=90000) as info:
                    page.locator(DOWNLOAD_LINK).first.click()
                download = info.value
                path = os.path.join(OUTPUT_DIR, f"{bu}_{cycle}_{download.suggested_filename}")
                download.save_as(path)
            except Exception as e:
                print(f"  no download ({type(e).__name__}: {str(e)[:200]})", flush=True)
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"{bu}_download_attempt.png"), full_page=True)
                continue
            with open(path, encoding="utf-8", errors="replace") as f:
                text = f.read()
            rows = list(csv.reader(io.StringIO(text)))
            print(f"  saved {os.path.basename(path)}: {len(rows)} rows", flush=True)
            for r in rows[:6]:
                print(f"    {r}", flush=True)
            hits = [r for r in rows if LNG_PATTERN.search(" ".join(r))]
            print(f"  {len(hits)} LNG-ish rows:", flush=True)
            for r in hits[:40]:
                print(f"    {r}", flush=True)
        browser.close()
    print(f"\nDONE. Outputs in {OUTPUT_DIR}/", flush=True)


if __name__ == "__main__":
    main()
