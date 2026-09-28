"""
Discovery pass over Eskom's public data portal: finds every CSV the
portal publishes (beyond the Station_Build_Up.csv the generation mix
script already uses), and prints each one's header, first/last rows and
row count, so a second archive tab (available capacity, planned/
unplanned outages, OCGT usage, EAF, ...) can be built against real
column names rather than guessed ones.

Three ways of finding files, since the portal's charts load their CSVs
from JavaScript rather than linking them plainly:
  1. WordPress's own media API (/dataportal/wp-json/wp/v2/media) - lists
     every uploaded file with its URL, if the API is left open.
  2. Crawl the portal's own pages (links under /dataportal/) and regex
     every "wp-content/uploads/....csv" string out of the raw HTML/JS.
  3. The WordPress pages API, to list portal pages step 2's link-following
     might miss.

Run by .github/workflows/south_africa_eskom_discovery.yml (manual
dispatch only) - eskom.co.za isn't reachable from every environment, but
is from GitHub's runners. Output goes to stdout and to
eskom_discovery_output.txt next to this script (uploaded as a workflow
artifact, not committed).
"""

import csv
import io
import os
import re
import sys
import time
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse

import requests

BASE = "https://www.eskom.co.za/dataportal/"
HEADERS = {"User-Agent": "gas-demand-scripts/1.0", "Cache-Control": "no-cache", "Pragma": "no-cache"}
MAX_PAGES = 60
# The first run sat for 12+ minutes with no output and had to be
# cancelled - Eskom's site can be very slow per request. So each request
# gets a (connect, read) timeout, the crawl stops at a time budget and
# moves on to describing whatever it has found, and every log line is
# written to the output file straight away so a killed run still leaves
# something to read.
TIMEOUT = (10, 20)
CRAWL_BUDGET_SECONDS = 6 * 60
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eskom_discovery_output.txt")

CSV_RE = re.compile(r"""https?:\\?/\\?/www\.eskom\.co\.za\\?/dataportal\\?/wp-content\\?/uploads\\?/[^"'\s<>)]+?\.csv""", re.I)
REL_CSV_RE = re.compile(r"""/dataportal/wp-content/uploads/[^"'\s<>)]+?\.csv""", re.I)
HREF_RE = re.compile(r"""href=["']([^"'#]+)["']""", re.I)

def log(text=""):
    print(text, flush=True)
    with open(OUTPUT_PATH, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def get(session, url, **kwargs):
    started = time.monotonic()
    try:
        r = session.get(url, timeout=TIMEOUT, **kwargs)
        elapsed = time.monotonic() - started
        if elapsed > 5:
            log(f"  (slow: {elapsed:.0f}s for {url})")
        return r
    except requests.RequestException as e:
        log(f"  ERROR {url}: {type(e).__name__}: {e}")
        return None


def media_api(session):
    found = {}
    for mime in ("text/csv", ""):
        page = 1
        while page <= 20:
            params = {"per_page": 100, "page": page}
            if mime:
                params["mime_type"] = mime
            r = get(session, urljoin(BASE, "wp-json/wp/v2/media"), params=params)
            if r is None or r.status_code != 200:
                log(f"  media API ({mime or 'all types'}) page {page}: HTTP {r.status_code if r is not None else 'n/a'}")
                break
            try:
                items = r.json()
            except ValueError:
                log(f"  media API page {page}: not JSON (API likely disabled)")
                break
            if not items:
                break
            for item in items:
                src = item.get("source_url", "")
                if src.lower().endswith((".csv", ".xlsx", ".xls")):
                    found[src] = item.get("title", {}).get("rendered", "")
            log(f"  media API ({mime or 'all types'}) page {page}: {len(items)} item(s)")
            if len(items) < 100:
                break
            page += 1
        if found:
            break
    return found


def pages_api(session):
    urls = set()
    page = 1
    while page <= 10:
        r = get(session, urljoin(BASE, "wp-json/wp/v2/pages"), params={"per_page": 100, "page": page, "_fields": "link,title"})
        if r is None or r.status_code != 200:
            log(f"  pages API page {page}: HTTP {r.status_code if r is not None else 'n/a'}")
            break
        try:
            items = r.json()
        except ValueError:
            log(f"  pages API page {page}: not JSON (API likely disabled)")
            break
        if not items:
            break
        for item in items:
            log(f"    page: {item.get('title', {}).get('rendered', '')!r} -> {item.get('link')}")
            urls.add(item.get("link"))
        if len(items) < 100:
            break
        page += 1
    return {u for u in urls if u}


def crawl(session, seeds):
    queue = [BASE] + sorted(seeds)
    seen = set()
    csvs = {}
    deadline = time.monotonic() + CRAWL_BUDGET_SECONDS
    while queue and len(seen) < MAX_PAGES:
        if time.monotonic() > deadline:
            log(f"  crawl time budget ({CRAWL_BUDGET_SECONDS}s) used up with {len(queue)} page(s) still queued - moving on")
            break
        url = queue.pop(0).split("#")[0]
        if url in seen:
            continue
        seen.add(url)
        r = get(session, url)
        if r is None or r.status_code != 200 or "html" not in r.headers.get("Content-Type", ""):
            continue
        html = r.text
        title = re.search(r"<title>(.*?)</title>", html, re.I | re.S)
        page_csvs = {m.replace("\\/", "/") for m in CSV_RE.findall(html)}
        page_csvs |= {urljoin("https://www.eskom.co.za", m) for m in REL_CSV_RE.findall(html)}
        log(f"  [{len(seen):>2}] {url}  ({title.group(1).strip() if title else ''}) -> {len(page_csvs)} csv ref(s)")
        for c in page_csvs:
            csvs.setdefault(c, url)
        for href in HREF_RE.findall(html):
            full = urljoin(url, href)
            p = urlparse(full)
            if p.netloc.endswith("eskom.co.za") and p.path.startswith("/dataportal/") and "wp-content" not in p.path \
                    and "wp-json" not in p.path and not p.path.endswith((".pdf", ".png", ".jpg", ".csv", ".xlsx")):
                if full not in seen:
                    queue.append(full)
    return csvs


def describe_csv(session, url):
    r = get(session, url, params={"t": int(datetime.now(timezone.utc).timestamp())})
    if r is None:
        return None
    log(f"\n=== {url}")
    log(f"    HTTP {r.status_code}, {len(r.content):,} bytes, Last-Modified: {r.headers.get('Last-Modified')}")
    if r.status_code != 200:
        return None
    rows = [row for row in csv.reader(io.StringIO(r.text)) if row]
    if not rows:
        log("    (empty)")
        return None
    log(f"    {len(rows) - 1} data row(s)")
    log(f"    HEADER: {rows[0]}")
    for row in rows[1:3]:
        log(f"    FIRST:  {row}")
    if len(rows) > 3:
        log(f"    LAST:   {rows[-1]}")
    return {
        "rows": len(rows) - 1,
        "header": rows[0],
        "first": rows[1][0] if len(rows) > 1 else None,
        "last": rows[-1][0] if len(rows) > 1 else None,
        "modified": r.headers.get("Last-Modified"),
    }


def main():
    open(OUTPUT_PATH, "w").close()
    session = requests.Session()
    session.headers.update(HEADERS)

    log("STEP 1 - WordPress media API")
    media = media_api(session)
    log(f"  -> {len(media)} data file(s) from media API")

    log("\nSTEP 2 - WordPress pages API")
    page_urls = pages_api(session)

    log("\nSTEP 3 - crawl portal pages for CSV references")
    crawled = crawl(session, page_urls)

    all_csvs = {}
    for url, title in media.items():
        all_csvs[url] = f"media API: {title}"
    for url, page in crawled.items():
        all_csvs.setdefault(url, f"referenced on {page}")
    for url, page in crawled.items():
        log(f"  crawl found: {url}  [on {page}]")

    # WordPress keeps every re-upload: the same dataset appears as
    # uploads/YYYY/MM/Name.csv, Name-1.csv ... Name-50.csv across many
    # months (the first full run printed hundreds of copies of the wind
    # load factor file alone). Group by base name with the "-N" suffix
    # stripped, and keep the copy in the latest YYYY/MM folder, preferring
    # the unsuffixed name there (the live file, as with Station_Build_Up).
    groups = {}
    for url in all_csvs:
        m = re.search(r"/uploads/(\d{4})/(\d{2})/([^/]+?)(?:-(\d+))?\.csv$", url, re.I)
        if not m:
            continue
        year, month, base, suffix = m.groups()
        key = (year, month, suffix is None, -int(suffix or 0))
        entry = groups.setdefault(base.lower(), {"count": 0, "best": None, "best_key": None, "base": base})
        entry["count"] += 1
        if entry["best_key"] is None or key > entry["best_key"]:
            entry["best"], entry["best_key"] = url, key

    log(f"\nSTEP 4 - {len(all_csvs)} CSV URL(s) found, {len(groups)} distinct dataset(s); newest copy of each:")
    summaries = []
    for name in sorted(groups):
        g = groups[name]
        summaries.append((g, describe_csv(session, g["best"])))

    log("\n" + "=" * 70)
    log(f"SUMMARY - {len(groups)} distinct Eskom data portal dataset(s)")
    log("=" * 70)
    for g, info in summaries:
        log(f"\n{g['base']}  ({g['count']} copies)")
        log(f"  {g['best']}")
        if info:
            log(f"  {info['rows']} rows, {info['first']} -> {info['last']}, Last-Modified {info['modified']}")
            log(f"  columns: {info['header']}")

    log(f"\nSaved: {OUTPUT_PATH}")
    if not groups:
        sys.exit(1)


if __name__ == "__main__":
    main()
