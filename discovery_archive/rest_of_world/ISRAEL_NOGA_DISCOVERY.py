"""
Follow-up to NIGERIA_ISRAEL_EGYPT_DISCOVERY.py (earlier this session):
Noga (Israel's Independent System Operator) homepage yielded real-
looking links but they were never explored further:
    /pdt/highvoltageproduction/
    /pdt/production-potential-in-renewable-energies/
    /statistics/annual-consumption-data/
    /statistics/annual-hourly-consumption-data/

Checking those pages directly (and guessing at an underlying API root
alongside them, the same way EPIAS/EirGrid's dashboards turned out to
have one) for real generation-mix data.
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 45)
BASE = "https://www.noga-iso.co.il"


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


PAGES = [
    "/en/pdt/highvoltageproduction/",
    "/en/pdt/production-potential-in-renewable-energies/",
    "/en/statistics/annual-consumption-data/",
    "/en/statistics/annual-hourly-consumption-data/",
]

for path in PAGES:
    r = try_get(f"page {path}", BASE + path)
    if r is not None and r.status_code == 200:
        text = r.text
        api_links = sorted(set(re.findall(r'(https?://[^\s"\'<>]*api[^\s"\'<>]*)', text, re.I)))[:10]
        json_links = sorted(set(re.findall(r'["\']([^"\']*\.json[^"\']*)["\']', text, re.I)))[:10]
        if api_links:
            print(f"  api-looking links: {api_links}", file=sys.stderr)
        if json_links:
            print(f"  json-looking links: {json_links}", file=sys.stderr)

try_get("API root guess 1", BASE + "/api/")
try_get("API root guess 2 (system data)", BASE + "/Umbraco/Api/Sysdata/GetSystemData")

# Found via web search: Noga runs a real Azure API Management developer
# portal - checking whether it lists any public (no-key) APIs.
r = try_get("Noga APIM developer portal", "https://apim-portal.noga-iso.co.il/")
if r is not None and r.status_code == 200:
    apis_links = sorted(set(re.findall(r'href="([^"]*api[^"]*)"', r.text, re.I)))[:15]
    print(f"  api-page links found: {apis_links}", file=sys.stderr)
