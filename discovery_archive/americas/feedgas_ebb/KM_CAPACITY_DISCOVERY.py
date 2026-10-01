"""
Second-pass discovery, following up on FEEDGAS_PIPES_DISCOVERY.py's
first real run (from a normal machine, not this sandbox - every URL
below came back 200 that time). That run only fetched each pipeline's
NOTICES/portal landing page; this one fetches the specific sub-pages
those landing pages link to that actually look like data, not menus.

Kinder Morgan's platform (pipeline2.kindermorgan.com) is the strongest
lead: the exact same URL pattern worked across GCX, NGPL, and Elba
Express (EEC) just by swapping `code=`, and its NGPL/EEC notices pages
both listed the same set of sub-pages - most promisingly:
  - /Capacity/OpAvailPoint.aspx?code=X - "Operationally Available"
    capacity BY POINT. OA capacity (subscribed minus scheduled, or
    similar - exact definition TBD from the real page) is adjacent to
    but not necessarily identical to nominated/scheduled quantity -
    this run will tell us which.
  - /LocationDataDownload/LocDataDwnld.aspx?code=X - named like a
    structured point-data export (CSV/Excel), not a report page -
    worth seeing what it actually returns.
  - /Capacity/OpAvailSegment.aspx?code=X and /Capacity/CapacityLog.aspx
    - checked too, lower priority (segment-level, or historical log
    rather than current point-level data).

Also re-checks Texas Eastern's already-reached "operationally
available" page more closely (rtba.enbridge.com) for its own internal
links, since the first pass's keyword scan may have missed report
links that don't literally contain the scanned keywords.

Covers three pipelines feeding three different terminals with ONE
platform's URL pattern: GCX (Corpus Christi LNG, Permian), NGPL
(Golden Pass LNG, Haynesville/LA side), Elba Express (Elba Island LNG).

Not runnable from the sandbox this repo is normally edited in - same
network block as every pipeline site tried so far. Run locally (same
venv as FEEDGAS_PIPES_DISCOVERY.py) and share the console output (or
the saved HTML in feedgas_discovery_html/) back.
"""

print("STARTING", flush=True)

import os
import re

import requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "km_capacity_discovery_html")

KM_BASE = "https://pipeline2.kindermorgan.com"
KM_CODES = [
    ("GCX", "Gulf Coast Express - feeds Corpus Christi LNG (Permian)"),
    ("NGPL", "NGPL - feeds Golden Pass LNG (Haynesville/LA side)"),
    ("EEC", "Elba Express - feeds Elba Island LNG"),
]
KM_PAGES = [
    ("Capacity/OpAvailPoint.aspx", "Operationally Available capacity by point"),
    ("LocationDataDownload/LocDataDwnld.aspx", "Location data download"),
    ("Capacity/OpAvailSegment.aspx", "Operationally Available capacity by segment"),
    ("Capacity/CapacityLog.aspx", "Capacity log"),
]

TARGETS = []
for page_path, page_desc in KM_PAGES:
    for code, code_desc in KM_CODES:
        label = f"Kinder Morgan {code} - {page_desc} ({code_desc})"
        TARGETS.append((label, f"{KM_BASE}/{page_path}?code={code}"))

# Re-check with a closer look, not just the first pass's keyword scan.
TARGETS.append((
    "Texas Eastern / TE (Enbridge) - operationally available (recheck)",
    "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu=TE&Type=OA",
))

DATA_KEYWORDS = [
    "scheduled quantity", "scheduled quantities", "nomination", "flow", "receipt", "delivery",
    "download", ".csv", ".xls", ".xlsx", "operationally available", "capacity", "point name",
]

JS_APP_MARKERS = [
    "you need to enable javascript", "please enable javascript", "id=\"root\"", "id=\"app\"",
    "ng-app", "data-reactroot", "__next",
]


def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30, verify=False)
        return r
    except requests.RequestException as e:
        print(f"  FAILED: {type(e).__name__}: {e}", flush=True)
        return None


def analyse(html):
    lower = html.lower()

    js_app_signals = [m for m in JS_APP_MARKERS if m in lower]
    if js_app_signals:
        print(f"  Looks like a JS-rendered single-page app (markers: {js_app_signals}) - plain requests may not see the real data.", flush=True)
    else:
        print("  Looks like a server-rendered page.", flush=True)

    found_keywords = sorted({k for k in DATA_KEYWORDS if k in lower})
    print(f"  Data-related keywords present: {found_keywords or 'none'}", flush=True)

    # A real HTML <table> with several rows is the strongest signal that
    # this page actually shows point-level data, not just a form/menu.
    table_count = lower.count("<table")
    row_count = lower.count("<tr")
    print(f"  <table> tags: {table_count}, <tr> tags: {row_count}", flush=True)

    # Direct download links (not javascript:__doPostBack, which needs a
    # simulated ASP.NET postback with real VIEWSTATE - out of scope for
    # a plain GET check like this one).
    links = re.findall(r'href=[\'"]([^\'" >]+)', html, flags=re.IGNORECASE)
    csv_xls_links = [l for l in links if re.search(r"\.(csv|xlsx?|txt)(\?|$)", l, re.IGNORECASE)]
    if csv_xls_links:
        print(f"  Direct file download link(s) found: {csv_xls_links[:10]}", flush=True)

    postback_count = sum(1 for l in links if l.lower().startswith("javascript:__dopostback"))
    if postback_count:
        print(f"  {postback_count} __doPostBack link(s) found (need simulated ASP.NET postback, not a plain GET).", flush=True)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for label, url in TARGETS:
        print(f"\n{'=' * 70}\n{label}\n{url}\n{'=' * 70}", flush=True)
        r = fetch(url)
        if r is None:
            continue

        print(f"  status: {r.status_code}, content-length: {len(r.content):,} bytes, content-type: {r.headers.get('Content-Type')}", flush=True)

        if r.status_code != 200:
            print(f"  Non-200 response - first 300 chars: {r.text[:300]!r}", flush=True)
            continue

        content_type = r.headers.get("Content-Type", "")
        if "text/html" not in content_type:
            # Might actually be a real CSV/Excel export, not HTML at all.
            safe_name = re.sub(r"[^a-zA-Z0-9]+", "_", label)[:60]
            ext = ".csv" if "csv" in content_type else (".xlsx" if "spreadsheet" in content_type or "excel" in content_type else ".bin")
            out_path = os.path.join(OUTPUT_DIR, safe_name + ext)
            with open(out_path, "wb") as f:
                f.write(r.content)
            print(f"  NON-HTML response (content-type: {content_type}) - saved raw -> {out_path}. This might be the actual data file.", flush=True)
            continue

        safe_name = re.sub(r"[^a-zA-Z0-9]+", "_", label)[:60] + ".html"
        out_path = os.path.join(OUTPUT_DIR, safe_name)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(r.text)
        print(f"  Saved -> {out_path}", flush=True)

        analyse(r.text)

    print(
        "\nDONE. Please share the console output above (or the saved files in "
        f"{OUTPUT_DIR}/) - especially anything with real <table>/<tr> content or "
        "a non-HTML content-type, since that's the actual data rather than a menu page.",
        flush=True,
    )


if __name__ == "__main__":
    main()
