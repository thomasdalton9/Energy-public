"""
Discovery pass for Jordan electricity generation data. NEPCO (National
Electric Power Company) is Jordan's grid operator; a web search found a
NEPCO dataset ("nepco-2170-2023") on Jordan's official open government
data portal, opendata.gov.jo - checking whether that portal exposes a
real API/CSV download (many government open-data portals run on CKAN,
which has a well-documented REST API at /api/3/action/...) rather than
just a landing page, and what the dataset actually contains.

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


def try_get(label, url, headers=None, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=headers or HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code}  bytes={len(r.content)}  content-type={r.headers.get('content-type')}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def probe_opendata_portal():
    for url in [
        "https://opendata.gov.jo/en/dataset/nepco-2170-2023",
        "https://www.opendata.gov.jo/dataset/nepco-2170-2023",
    ]:
        r = try_get("opendata.gov.jo NEPCO dataset page", url)
        if r is not None and r.status_code == 200:
            print(r.text[:3000])
            resource_links = sorted(set(re.findall(r'href=["\']([^"\']*(?:\.csv|\.xlsx?|/resource/|/download/)[^"\']*)["\']', r.text, re.I)))
            print(f"\n  {len(resource_links)} resource/download links found:")
            for l in resource_links:
                print(f"    {l}")

    # CKAN's own REST API, if this portal runs CKAN (common for gov
    # open-data sites) - would give clean JSON with every resource URL
    # directly, no HTML scraping needed
    for url in [
        "https://opendata.gov.jo/api/3/action/package_show?id=nepco-2170-2023",
        "https://opendata.gov.jo/en/api/3/action/package_show?id=nepco-2170-2023",
        "https://opendata.gov.jo/api/3/action/package_search?q=nepco",
    ]:
        r = try_get("opendata.gov.jo CKAN API guess", url)
        if r is not None and r.status_code == 200:
            print(r.text[:3000])


def probe_nepco_site():
    for url in [
        "https://nepco.com.jo/en/",
        "https://nepco.com.jo/en/electrical_energy_en.aspx",
    ]:
        r = try_get("nepco.com.jo", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2000])


def main():
    probe_opendata_portal()
    probe_nepco_site()


if __name__ == "__main__":
    main()
