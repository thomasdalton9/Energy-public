"""
Discovery pass for New Zealand electricity market data, prompted by a
user-shared link to app.em6.co.nz (Enerlytica's "EM6 Live" dashboard -
stacked generation-weighted-average-price by grid zone, wind forecast).
That's a commercial/subscription platform's own front-end, not
obviously a public data source - checking here whether it (or its
underlying API) is usable without a login, and in parallel checking
NZ's actual official public sources:
  - EMI (Electricity Market Information), emi.ea.govt.nz - the
    Electricity Authority's own public market data repository
    (generation, pricing, demand), analogous to AEMO/EIA/ENTSO-E for
    NZ. This is the durable, official, no-login source if it has a
    usable download/API path.
  - Transpower (the grid operator)'s own real-time generation dataset,
    published separately from EMI.

This script just probes reachability/structure - no assumptions
carried in from outside this run.
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


def probe_em6():
    # app.em6.co.nz is an Angular/React-style SPA (params in the URL
    # fragment/query control chart filters client-side) - the HTML
    # itself won't have the data, but checking whether it's reachable
    # at all and whether an obvious public API sits behind it (common
    # pattern: <app>.<vendor>.co.nz talks to api.<vendor>.co.nz or
    # a /api/ path on the same host).
    for url in [
        "https://app.em6.co.nz/",
        "https://api.em6.co.nz/",
        "https://em6.co.nz/",
    ]:
        r = try_get("em6.co.nz", url)
        if r is not None and r.status_code == 200 and "text/html" in (r.headers.get("content-type") or ""):
            print(r.text[:2000])


def probe_emi():
    # EMI (Electricity Authority's Electricity Market Information) -
    # NZ's official public market data site. Checking the main portal
    # and its known "wholesale market" data export area.
    for url in [
        "https://www.emi.ea.govt.nz/",
        "https://www.emi.ea.govt.nz/Wholesale/Datasets",
        "https://www.emi.ea.govt.nz/Wholesale/Reports",
    ]:
        r = try_get("emi.ea.govt.nz", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2500])


def probe_transpower():
    # Transpower (NZ grid operator) publishes real-time generation
    # data separately - checking their main data portal.
    for url in [
        "https://www.transpower.co.nz/system-operator/live-system-and-market-data",
        "https://www.transpower.co.nz/",
    ]:
        r = try_get("transpower.co.nz", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2000])


def main():
    probe_em6()
    probe_emi()
    probe_transpower()


if __name__ == "__main__":
    main()
