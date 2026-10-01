"""
One-off discovery probe for GIIGNL (Groupe International des Importateurs
de Gaz Naturel Liquefie / International Group of LNG Importers) - checks
whether their annual "LNG Industry" report is available as an actual
downloadable file (PDF/Excel/CSV) or structured data, versus only a
narrative report page - same question already asked (and answered "no,
narrative only") for EIA's coal markets page.

Fetches the GIIGNL site's publications/reports page and looks for direct
file download links, an API-like path, and any indication of
machine-readable data (not yet confirmed against the live page).
"""

import re
import sys

import requests

CANDIDATE_URLS = [
    "https://giignl.org/publications/",
    "https://giignl.org/the-lng-industry/",
    "https://giignl.org/",
]
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)


def probe(url):
    print(f"\n{'=' * 70}\nFetching {url} ...", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
    except requests.RequestException as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
        return
    print(f"  status {r.status_code}, final URL {r.url}, {len(r.content):,} bytes", file=sys.stderr)
    if r.status_code != 200:
        return
    html = r.text

    links = re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.IGNORECASE)
    file_links = [l for l in links if re.search(r"\.(pdf|xlsx?|csv|json)(\?|$)", l, re.IGNORECASE)]
    print(f"  direct file links ({len(file_links)}):")
    for l in sorted(set(file_links))[:20]:
        print(f"    {l}")

    report_links = [l for l in links if re.search(r"report|lng.?industry|annual|publication", l, re.IGNORECASE)]
    print(f"  report-like links ({len(report_links)}):")
    for l in sorted(set(report_links))[:20]:
        print(f"    {l}")

    print(f"  <table count: {html.lower().count('<table')}, <tr count: {html.lower().count('<tr')}")


def main():
    for url in CANDIDATE_URLS:
        probe(url)


if __name__ == "__main__":
    main()
