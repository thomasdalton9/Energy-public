"""
Probe CGES (Crnogorski elektroprenosni sistem, Montenegro's TSO) for open
generation / load / transparency data. Bounded crawl: the root pages in
English and Montenegrin, then internal links whose text or URL mentions
generation, production, load, balance, transparency, report or a data file.

Lists data-file links (xls/xlsx/csv/json/xml/pdf), JSON/API-looking URLs
referenced by page scripts, and tables seen on pages, so we can see whether
there is a machine-readable feed and how far back it goes.

Usage: python3 CGES_PROBE.py   (writes cges_probe_report.json)
"""
import json
import re
import sys
import time
from urllib.parse import urljoin, urlparse

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en,me;q=0.8",
}
TIMEOUT = (10, 45)
ROOTS = ["https://www.cges.me/", "https://www.cges.me/en/", "https://cges.me/", "https://www.cges.me/en/home"]
KEYWORDS = re.compile(r"generat|product|proizvod|load|optere|balanc|bilans|transpar|report|izvje|izvješ|data|podac|"
                      r"download|market|trziste|tržište|elektroenerg|system|sistem|statist|exchange|razmjen", re.I)
FILE_EXT = re.compile(r"\.(xlsx?|csv|json|xml|zip|pdf)(\?|$)", re.I)
MAX_PAGES = 45
seen, files, apis, tables, errors = {}, set(), set(), [], []


def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        return r
    except requests.RequestException as e:
        errors.append((url, f"{type(e).__name__}: {str(e)[:120]}"))
        return None


def page_info(url, r):
    text = r.text
    title = re.search(r"<title[^>]*>(.*?)</title>", text, re.S | re.I)
    hrefs = {urljoin(r.url, h.split("#")[0]) for h in re.findall(r'href=["\']([^"\']+)["\']', text) if not h.startswith(("mailto:", "javascript:", "tel:"))}
    for h in hrefs:
        if FILE_EXT.search(h):
            files.add(h)
    for m in re.findall(r'["\'](https?://[^"\']+|/[\w\-./]+)["\']', text):
        if re.search(r"(api|json|ajax|chart|graph|data)", m, re.I) and not FILE_EXT.search(m) and len(m) < 200 and not re.search(r"\.(css|png|jpe?g|gif|svg|woff2?|ico)", m, re.I):
            apis.add(urljoin(r.url, m))
    n_tables = len(re.findall(r"<table", text, re.I))
    if n_tables:
        tables.append((r.url, n_tables, len(re.findall(r"<tr", text, re.I))))
    return (re.sub(r"\s+", " ", title.group(1)).strip()[:90] if title else None), hrefs


def main():
    queue, host = [], None
    for root in ROOTS:
        r = fetch(root)
        print(f"ROOT {root}: {None if r is None else r.status_code} {'' if r is None else r.url} {'' if r is None else len(r.content)}B", flush=True)
        if r is not None and r.status_code == 200 and host is None:
            host = urlparse(r.url).netloc
            title, hrefs = page_info(root, r)
            seen[r.url] = (r.status_code, title)
            queue += sorted(h for h in hrefs if urlparse(h).netloc == host and KEYWORDS.search(h))
    if host is None:
        print("No CGES root reachable; errors:", errors)
        return
    i = 0
    while queue and len(seen) < MAX_PAGES:
        url = queue.pop(0)
        if url in seen or FILE_EXT.search(url) or url.rstrip("/") + "/" in seen:
            continue
        r = fetch(url)
        time.sleep(0.5)
        if r is None:
            seen[url] = (None, None)
            continue
        title, hrefs = page_info(url, r)
        seen[url] = (r.status_code, title)
        print(f"  [{r.status_code}] {url} | {title}", flush=True)
        if r.status_code == 200 and "html" in r.headers.get("Content-Type", ""):
            for h in sorted(hrefs):
                if urlparse(h).netloc == host and h not in seen and h not in queue and KEYWORDS.search(h):
                    queue.append(h)
    print(f"\nPAGES crawled: {len(seen)}")
    print(f"\nDATA FILES ({len(files)}):")
    for f in sorted(files)[:80]:
        print("  ", f)
    print(f"\nAPI/JSON-looking URLs in scripts ({len(apis)}):")
    for a in sorted(apis)[:60]:
        print("  ", a)
    print(f"\nPAGES WITH TABLES ({len(tables)}):")
    for t in tables[:30]:
        print("  ", t)
    print(f"\nERRORS ({len(errors)}):", errors[:10])
    # sample the newest-looking data files: HEAD for size / last-modified
    print("\nFILE HEADERS (first 15):")
    for f in sorted(files)[:15]:
        try:
            h = requests.head(f, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
            print(f"  {h.status_code} {h.headers.get('Content-Length')}B lm={h.headers.get('Last-Modified')} {f}")
        except requests.RequestException as e:
            print(f"  ERR {f} {type(e).__name__}")
    with open("cges_probe_report.json", "w") as fh:
        json.dump({"pages": seen, "files": sorted(files), "apis": sorted(apis), "tables": tables, "errors": errors}, fh, indent=1, default=str)


if __name__ == "__main__":
    main()
