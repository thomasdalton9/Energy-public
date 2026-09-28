"""
Third-pass discovery for Kinder Morgan's "Operationally Available
capacity by point" page (OpAvailPoint.aspx). The saved HTML from the
second pass revealed why a plain `requests.get()` never sees any data:
the grid (id="...DGOpAvail") is an Infragistics WebDataGrid that ships
as an EMPTY shell in the initial server response - a self-closing
<tbody data-ig="...mkr:rows" .../> with zero rows, visibility:hidden -
and only gets populated client-side, in the browser, after clicking
the "Retrieve Data" button (id="...HeaderBTN1_btnRetrieve"). No amount
of HTML parsing on the initial GET response will ever find the actual
capacity numbers; a real browser has to click the button and let its
JavaScript run.

The form's default values already look sensible with no changes
needed: "Eff Gas Day" defaults to today's date, "Cycle Selection"
defaults to "BEST AVAILABLE", and the location radio defaults to
"Delivery Location/Delivery point(s) quantity" - so this just clicks
Retrieve with the defaults as-is, rather than trying to also automate
picking different query criteria.

There's also a "Download Data" button (id="...HeaderBTN1_btnDownload")
right next to Retrieve - this script tries that too, since an actual
file export (if it works, and if it doesn't require Retrieve to have
run first) would be far more reliable to parse than scraping rendered
grid HTML, which may involve virtualized/windowed rows that don't all
exist in the DOM at once.

Needs Playwright (not currently confirmed installed in the venv this
has been run from - `pip install playwright` then
`playwright install chromium` if the import below fails).

Not runnable from the sandbox this repo is normally edited in - same
network block as every pipeline site tried so far. Run locally and
share back: the console output, the saved screenshots, the saved grid
HTML dump, and - if the download button produces one - the actual
downloaded file.
"""

import os
import sys

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

RETRIEVE_BTN = "#WebSplitter1_tmpl1_ContentPlaceHolder1_HeaderBTN1_btnRetrieve"
DOWNLOAD_BTN = "#WebSplitter1_tmpl1_ContentPlaceHolder1_HeaderBTN1_btnDownload"
GRID_CONTAINER = "#WebSplitter1_tmpl1_ContentPlaceHolder1_DGOpAvail"


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as p:
        # ignore_https_errors belongs on the page/context, not launch() -
        # launch() rejects it with a TypeError before anything runs.
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(ignore_https_errors=True)

        print(f"Navigating to {URL} ...", flush=True)
        page.goto(URL, wait_until="networkidle", timeout=30000)
        page.screenshot(path=os.path.join(OUTPUT_DIR, "01_initial_load.png"), full_page=True)
        print("Saved 01_initial_load.png", flush=True)

        # ---- try Retrieve Data ----
        print("\nClicking 'Retrieve Data' with default query values (today's Eff Gas Day, BEST AVAILABLE cycle, Delivery locations)...", flush=True)
        try:
            page.click(RETRIEVE_BTN, timeout=10000)
            page.wait_for_load_state("networkidle", timeout=30000)
            page.wait_for_timeout(2000)  # let any client-side grid rendering settle
        except Exception as e:
            print(f"  Click/wait FAILED: {type(e).__name__}: {e}", flush=True)

        page.screenshot(path=os.path.join(OUTPUT_DIR, "02_after_retrieve.png"), full_page=True)
        print("Saved 02_after_retrieve.png - LOOK AT THIS ONE FIRST.", flush=True)

        try:
            grid_html = page.inner_html(GRID_CONTAINER, timeout=5000)
            grid_text = page.inner_text(GRID_CONTAINER, timeout=5000)
            with open(os.path.join(OUTPUT_DIR, "grid_after_retrieve.html"), "w", encoding="utf-8") as f:
                f.write(grid_html)
            with open(os.path.join(OUTPUT_DIR, "grid_after_retrieve.txt"), "w", encoding="utf-8") as f:
                f.write(grid_text)
            print(f"Saved grid_after_retrieve.html/.txt - grid visible text is {len(grid_text)} chars", flush=True)
            print(f"First 500 chars of grid text:\n{grid_text[:500]}", flush=True)
        except Exception as e:
            print(f"  Could not read grid container: {type(e).__name__}: {e}", flush=True)

        # Full page HTML after retrieve, in case the grid ended up under a
        # different container than expected.
        with open(os.path.join(OUTPUT_DIR, "full_page_after_retrieve.html"), "w", encoding="utf-8") as f:
            f.write(page.content())
        print("Saved full_page_after_retrieve.html", flush=True)

        # ---- try Download Data ----
        print("\nClicking 'Download Data'...", flush=True)
        try:
            with page.expect_download(timeout=10000) as download_info:
                page.click(DOWNLOAD_BTN, timeout=10000)
            download = download_info.value
            save_path = os.path.join(OUTPUT_DIR, "download_" + download.suggested_filename)
            download.save_as(save_path)
            print(f"  A real file download happened! Saved -> {save_path}", flush=True)
        except Exception as e:
            print(f"  No direct download captured ({type(e).__name__}: {e}) - probably opened a popup/dialog instead. Screenshotting current state.", flush=True)
            page.wait_for_timeout(1500)
            page.screenshot(path=os.path.join(OUTPUT_DIR, "03_after_download_click.png"), full_page=True)
            print("  Saved 03_after_download_click.png - check if a format-selection dialog appeared.", flush=True)
            # Also dump all open pages/popups in case it opened a new tab/window.
            for i, pg in enumerate(browser.contexts[0].pages):
                if pg is not page:
                    pg.screenshot(path=os.path.join(OUTPUT_DIR, f"03b_popup_{i}.png"), full_page=True)
                    print(f"  Saved 03b_popup_{i}.png (a popup/new tab opened)", flush=True)

        browser.close()

    print(
        f"\nDONE. Please share everything in {OUTPUT_DIR}/ back - especially "
        "02_after_retrieve.png (does the grid show real rows now?), "
        "grid_after_retrieve.txt (the actual text, if any), and anything from "
        "the Download Data attempt.",
        flush=True,
    )


if __name__ == "__main__":
    main()
