"""
Discovery pass for Cyprus electricity generation data. Cyprus runs an
isolated island grid (not interconnected to the European mainland), so
it is NOT covered by ENTSO-E's Transparency Platform the way every
other EU country in this repo is (see europe/entsoe_powergen_monthly_2022_2026.py) -
a genuinely separate source is needed.

Candidate found via web search: CyCEM Insights (cycem-insights.cyi.ac.cy),
a free platform launched by the Cyprus Institute presenting the
Cypriot Competitive Electricity Market's data - prices, trading
volumes, generation mix, RES share, system balancing - with
"structured datasets and export functions". The official TSO
(Transmission System Operator Cyprus, tsoc.org.cy) is the other
candidate, in case it publishes raw data directly rather than only
through CyCEM.

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


def probe_cycem_insights():
    r = try_get("CyCEM Insights homepage", "https://cycem-insights.cyi.ac.cy/")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])
        # look for API/data endpoint hints in the page's own JS
        links = sorted(set(re.findall(r'["\']((?:https?:)?//[^"\']*(?:api|data)[^"\']*)["\']', r.text, re.I)))
        print(f"\n  {len(links)} api/data-like URLs referenced in the page:")
        for l in links[:30]:
            print(f"    {l}")
        script_srcs = sorted(set(re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', r.text)))
        print(f"\n  {len(script_srcs)} external script files referenced:")
        for s in script_srcs[:20]:
            print(f"    {s}")


def probe_tsoc():
    for url in [
        "https://tsoc.org.cy/",
        "https://tsoc.org.cy/en/",
        "https://tsoc.org.cy/electrical-system/system-data/",
    ]:
        r = try_get("tsoc.org.cy", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2000])


def main():
    probe_cycem_insights()
    probe_tsoc()


if __name__ == "__main__":
    main()
