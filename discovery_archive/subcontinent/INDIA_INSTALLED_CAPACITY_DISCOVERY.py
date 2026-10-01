"""
Discovery pass for India's "All India Installed Capacity of Power
Stations" - CEA's well-known monthly report giving installed capacity
(MW) broken down by region/state/sector (Central/State/Private) and
fuel/mode (Thermal-Coal/Lignite/Gas/Diesel, Nuclear, Hydro, RES -
Wind/Solar/Biomass/Small-Hydro). Sibling to india_coal_stock.py, which
already confirmed npp.gov.in (CEA's National Power Portal) is reachable
from GitHub Actions runners with a predictable date-based report URL -
checking whether the same public-reports tree (or CEA's own cea.nic.in
site) publishes this report too.

This script just probes candidate paths and reports what's actually
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


def probe_cea_nic_in():
    # CEA's own site - the report's well-known public home
    # ("cea.nic.in > Reports > Generation > Installed Capacity Reports").
    # Guessed landing paths first; if any resolve, their HTML should
    # reveal the real monthly-PDF link pattern directly.
    for url in [
        "https://cea.nic.in/installed-capacity-report/",
        "https://cea.nic.in/installed-capacity-report",
        "https://cea.nic.in/wp-content/uploads/installed_capacity/2026/09/IC_Sep_2026.pdf",
        "https://cea.nic.in/",
    ]:
        r = try_get("cea.nic.in", url)
        if r is not None and r.status_code == 200 and "text/html" in (r.headers.get("content-type") or ""):
            print(r.text[:3000])


def probe_npp_public_reports():
    # npp.gov.in's public-reports tree already confirmed reachable
    # (india_coal_stock.py) with predictable date-based subpaths under
    # public-reports/cea/... - checking sibling trees for a monthly
    # capacity report, and a listing page if one exists.
    for url in [
        "https://npp.gov.in/public-reports/cea/monthly",
        "https://npp.gov.in/public-reports/cea/installed-capacity",
        "https://npp.gov.in/publicreports/CEA/monthly/generation/",
        "https://npp.gov.in/public-reports",
    ]:
        r = try_get("npp.gov.in public-reports tree", url)
        if r is not None and r.status_code == 200:
            print(r.text[:3000])


def probe_npp_dashboard_api():
    # npp.gov.in also runs live dashboards (india_powergen_npp.py's
    # gc-map-dashboard-meritchart); checking whether an "installed
    # capacity" dashboard/API exists under the same host as a JSON
    # alternative to a monthly PDF.
    for url in [
        "https://npp.gov.in/dashBoard/cp-map-dashboard",
        "https://npp.gov.in/dashboard/installed-capacity",
        "https://npp.gov.in/InstalledCapacity",
    ]:
        r = try_get("npp.gov.in dashboard guess", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2000])


def main():
    probe_cea_nic_in()
    probe_npp_public_reports()
    probe_npp_dashboard_api()


if __name__ == "__main__":
    main()
