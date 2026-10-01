"""
Discovery pass for India hydro reservoir storage ("hydro stocks") -
sibling to india_coal_stock.py (CEA daily coal stock report), same
idea for the water-storage side of Indian power generation.

Candidate source: the Central Water Commission (CWC) publishes a
weekly "Reservoir Storage Bulletin" covering ~150+ monitored reservoirs
nationally, widely cited by press/analysts for hydropower generation
context. Public site is cwc.gov.in; India-WRIS (indiawris.gov.in) also
aggregates the same underlying CWC data via its own portal/API.

This script just probes both paths and reports what's actually
reachable/parseable - no assumptions carried in from outside this run.
"""

import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code}  bytes={len(r.content)}  content-type={r.headers.get('content-type')}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def probe_cwc():
    r = try_get("CWC main site", "https://cwc.gov.in/")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])

    # CWC's reservoir storage bulletin - guessed landing page path,
    # likely wrong first try but the real navigation should be visible
    # in the main site's HTML above regardless.
    r = try_get("CWC reservoir storage bulletin (guessed path)",
                "https://cwc.gov.in/reservoir-storage-bulletin")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])


def probe_india_wris():
    r = try_get("India-WRIS main site", "https://indiawris.gov.in/")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])


def main():
    probe_cwc()
    probe_india_wris()


if __name__ == "__main__":
    main()
