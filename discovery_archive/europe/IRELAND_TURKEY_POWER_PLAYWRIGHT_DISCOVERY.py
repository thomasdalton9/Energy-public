"""
Follow-up to IRELAND_TURKEY_DISCOVERY.py.

Ireland gas is solved (data.gov.ie's CKAN API gave direct JSON/CSV URLs
for Gas Networks Ireland's Daily Gas Demand dataset - no browser needed).

Ireland power (smartgriddashboard.com) and Turkey power
(seffaflik.epias.com.tr) both returned plain HTML shells to a bare
requests.get() with no obvious API links - both are very likely
client-rendered dashboards (same situation Sri Lanka's PUCSL dashboard
was in, solved in SRILANKA_PUCSL_PLAYWRIGHT_DISCOVERY.py). Reusing that
exact Playwright network-capture approach here: load the real page,
record every XHR/fetch response while it renders and after clicking
likely controls, and look for the underlying JSON API.
"""

import sys

from playwright.sync_api import sync_playwright

INTERESTING_EXT = (".json", ".csv", ".xlsx", ".xls")

TARGETS = [
    ("Ireland - Smart Grid Dashboard", "https://www.smartgriddashboard.com/",
     ["Generation", "Demand", "Wind", "CO2", "Interconnection", "Fuel Mix", "System"]),
    ("Turkey - EPIAS Transparency Platform", "https://seffaflik.epias.com.tr/",
     ["Üretim", "Generation", "Gerçek Zamanlı", "Tüketim", "Doğal Gaz", "Elektrik"]),
]


def probe_page(playwright, label, url, click_labels):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    browser = playwright.chromium.launch()
    page = browser.new_page()
    captured = []

    def on_response(response):
        req = response.request
        ctype = response.headers.get("content-type", "")
        is_interesting = (
            req.resource_type in ("xhr", "fetch")
            or any(req.url.lower().split("?")[0].endswith(ext) for ext in INTERESTING_EXT)
            or "json" in ctype
        )
        if is_interesting:
            captured.append((req.method, req.url, response.status, ctype))

    page.on("response", on_response)

    try:
        page.goto(url, wait_until="networkidle", timeout=45000)
    except Exception as e:
        print(f"  navigation error (continuing anyway): {type(e).__name__}: {e}")
    page.wait_for_timeout(6000)

    print(f"  {len(captured)} XHR/fetch/json responses captured after load:")
    for method, req_url, status, ctype in captured:
        print(f"    {method} {status} {ctype} -> {req_url}")

    for label_click in click_labels:
        try:
            locator = page.get_by_text(label_click, exact=False)
            if locator.count() > 0:
                print(f"  clicking control: {label_click!r}")
                locator.first.click(timeout=5000)
                page.wait_for_timeout(4000)
        except Exception as e:
            print(f"  couldn't click {label_click!r}: {type(e).__name__}: {e}")

    print(f"  {len(captured)} total XHR/fetch/json responses after interaction:")
    for method, req_url, status, ctype in captured:
        print(f"    {method} {status} {ctype} -> {req_url}")

    body_text = page.inner_text("body")
    print(f"  visible body text length: {len(body_text)} chars")
    print(f"  first 1500 chars: {body_text[:1500]}")

    browser.close()


def main():
    with sync_playwright() as playwright:
        for label, url, click_labels in TARGETS:
            try:
                probe_page(playwright, label, url, click_labels)
            except Exception as e:
                print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
