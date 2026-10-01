"""
Follow-up to CARIBBEAN_SRILANKA_DISCOVERY.py: gendata.pucsl.gov.lk/home
is a real, live page (confirmed 200 OK) but it's a Next.js
client-rendered app - its actual dispatch/generation data loads via
JavaScript after the page renders, so a plain requests.get() only sees
the empty shell HTML, not the data itself.

Using Playwright (same approach that found Bolivia's CNDC API in
BOLIVIA_CNDC_PLAYWRIGHT_DISCOVERY.py) to load the page for real and
capture every XHR/fetch request it makes while rendering, looking for
the underlying JSON API that actually carries the generation dispatch
data - once found, that endpoint can likely be called directly with
plain requests, without needing a browser on every real run.
"""

import sys

from playwright.sync_api import sync_playwright

PAGES = [
    "https://gendata.pucsl.gov.lk/home",
]

INTERESTING_EXT = (".json", ".csv", ".xlsx", ".xls")


def probe_page(playwright, url):
    print(f"\n{'=' * 70}\n{url}\n{'=' * 70}", file=sys.stderr)
    browser = playwright.chromium.launch()
    page = browser.new_page()
    captured = []

    def on_response(response):
        req = response.request
        ctype = response.headers.get("content-type", "")
        is_interesting = (
            req.resource_type in ("xhr", "fetch")
            or any(req.url.lower().endswith(ext) for ext in INTERESTING_EXT)
            or "json" in ctype or "csv" in ctype
        )
        if is_interesting:
            captured.append((req.method, req.url, response.status, ctype))

    page.on("response", on_response)

    try:
        page.goto(url, wait_until="networkidle", timeout=45000)
    except Exception as e:
        print(f"  navigation error (continuing anyway): {type(e).__name__}: {e}")
    page.wait_for_timeout(6000)

    print(f"  {len(captured)} XHR/fetch/data responses captured:")
    for method, req_url, status, ctype in captured:
        print(f"    {method} {status} {ctype} -> {req_url}")

    # try clicking any date-range / dropdown controls that might reveal
    # further endpoints (exact labels unknown - try common ones)
    for label in ["Today", "Daily", "Generation", "Dispatch", "Search", "View"]:
        try:
            locator = page.get_by_text(label, exact=False)
            if locator.count() > 0:
                print(f"  clicking control: {label!r}")
                locator.first.click(timeout=5000)
                page.wait_for_timeout(4000)
        except Exception as e:
            print(f"  couldn't click {label!r}: {type(e).__name__}: {e}")

    print(f"  {len(captured)} total XHR/fetch/data responses after interaction:")
    for method, req_url, status, ctype in captured:
        print(f"    {method} {status} {ctype} -> {req_url}")

    body_text = page.inner_text("body")
    print(f"  visible body text length: {len(body_text)} chars")
    print(f"  first 2000 chars: {body_text[:2000]}")

    browser.close()


def main():
    with sync_playwright() as playwright:
        for url in PAGES:
            probe_page(playwright, url)


if __name__ == "__main__":
    main()
