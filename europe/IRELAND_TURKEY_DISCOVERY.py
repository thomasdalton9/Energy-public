"""
Discovery for Ireland (power + gas) and Turkey (power + gas).

Ireland gas: Gas Networks Ireland publishes real open data - confirmed
via web search - "Daily Gas Demand Ireland" (data.gov.ie), split into
NDM (residential/SME), DM/LDM (large industrial), and Power Generation,
daily since 2018-01-01, quarterly CSV files at a known URL pattern
(gasnetworks.ie/corporate/open-data/{year}-Q{n}-Daily-Gas-Demand.csv).
Checking the data.gov.ie dataset page for the real resource list (exact
quarterly URLs), since guessing the current quarter's filename directly
is fragile.

Ireland power: EirGrid/SONI publish an all-island fuel mix, and SONI's
"Smart Grid Dashboard" reportedly supports CSV download - checking for
its real underlying API (smartgriddashboard.com is EirGrid Group's own
site for this, based on general knowledge of the Irish grid).

Turkey power: EPIAS Transparency Platform (seffaflik.epias.com.tr) -
confirmed via web search to have real-time generation by resource type,
hourly, since 2015. Checking its actual API root for a no-key public
endpoint (some transparency platforms have public read endpoints
alongside authenticated ones).

Turkey gas: checking EPIAS's transparency platform for a gas section
too (many exchange operators, like EPIAS, cover both power and gas) and
looking for a BOTAS or EPDK open data page.

Usage: python3 IRELAND_TURKEY_DISCOVERY.py
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 45)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
              file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def ireland_gas():
    r = try_get("data.gov.ie Daily Gas Demand dataset (CKAN API)",
                "https://data.gov.ie/api/3/action/package_show", params={"id": "dailygasdemandireland"})
    if r is not None and r.status_code == 200:
        try:
            data = r.json()
            resources = data.get("result", {}).get("resources", [])
            print(f"  {len(resources)} resources", file=sys.stderr)
            for res in resources[:10]:
                print(f"    {res.get('name')}: format={res.get('format')} url={res.get('url')}", file=sys.stderr)
        except ValueError:
            print(r.text[:1000], file=sys.stderr)


def ireland_power():
    for url in ["https://www.smartgriddashboard.com/api/", "https://www.smartgriddashboard.com/",
                "https://www.eirgridgroup.com/how-the-grid-works/", "https://smartgriddashboard.eirgrid.com/"]:
        r = try_get("Ireland power candidate", url, allow_redirects=True)
        if r is not None and r.status_code == 200:
            text = r.text
            links = sorted(set(re.findall(r'(https?://[^\s"\'<>]+api[^\s"\'<>]*)', text, re.I)))[:10]
            print(f"  api-looking links: {links}", file=sys.stderr)


def turkey_power():
    r = try_get("EPIAS transparency platform homepage", "https://seffaflik.epias.com.tr/")
    if r is not None and r.status_code == 200:
        text = r.text
        links = sorted(set(re.findall(r'(https?://[^\s"\'<>]+(?:api|servis)[^\s"\'<>]*)', text, re.I)))[:15]
        print(f"  api/servis-looking links: {links}", file=sys.stderr)
    r2 = try_get("EPIAS public API root guess", "https://seffaflik.epias.com.tr/electricity-service/v1/generation/data/realtime-generation")
    if r2 is not None:
        print(f"  body preview: {r2.text[:500]}", file=sys.stderr)


def turkey_gas():
    r = try_get("EPIAS gas section check", "https://seffaflik.epias.com.tr/", params={})
    if r is not None and r.status_code == 200:
        hits = [k for k in ["gas", "dogalgaz", "doğalgaz", "natural gas"] if k in r.text.lower()]
        print(f"  gas keyword hits on EPIAS homepage: {hits}", file=sys.stderr)
    for url in ["https://rapor.epias.com.tr/", "https://www.botas.gov.tr/"]:
        try_get("Turkey gas candidate", url)


if __name__ == "__main__":
    for label, fn in [("Ireland gas", ireland_gas), ("Ireland power", ireland_power),
                       ("Turkey power", turkey_power), ("Turkey gas", turkey_gas)]:
        print(f"\n=== {label} ===", file=sys.stderr)
        try:
            fn()
        except Exception as e:
            print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
