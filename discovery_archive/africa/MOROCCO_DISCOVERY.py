"""
Discovery pass for Morocco electricity generation data. ONEE (Office
National de l'Electricite et de l'Eau Potable) is Morocco's state grid
operator/utility; no dedicated ONEE open-data portal turned up in a web
search. Checking Morocco's national open-data portal (data.gov.ma) and
MASEN (Moroccan Agency for Solar Energy, which also publishes some
national generation-mix figures as part of its renewables reporting).

This script just probes reachability/structure - no assumptions
carried in from outside this run.
"""

import re
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


def probe_data_gov_ma():
    r = try_get("data.gov.ma homepage", "https://www.data.gov.ma/")
    if r is not None and r.status_code == 200:
        print(r.text[:2000])

    # if this portal is also CKAN (common for gov open-data sites),
    # its search API would find ONEE/electricity datasets directly
    for url in [
        "https://www.data.gov.ma/api/3/action/package_search?q=electricite",
        "https://www.data.gov.ma/api/3/action/package_search?q=ONEE",
        "https://www.data.gov.ma/api/3/action/package_search?q=electricity",
    ]:
        r = try_get("data.gov.ma CKAN API guess", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2000])


def probe_onee_masen():
    for url in [
        "https://www.one.org.ma/",
        "https://www.masen.ma/en",
    ]:
        r = try_get("ONEE/MASEN", url)
        if r is not None and r.status_code == 200:
            print(r.text[:1500])


def main():
    probe_data_gov_ma()
    probe_onee_masen()


if __name__ == "__main__":
    main()
