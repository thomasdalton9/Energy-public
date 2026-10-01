"""
Discovery pass for China's National Bureau of Statistics (NBS) as an
energy-data source - specifically the monthly headline indicators NOT
already covered by Ember (asia/ember_china_generation_capacity.py only
has electricity generation/capacity by fuel, not primary-fuel
production): raw coal output, crude oil output, natural gas output,
and total energy consumption.

Two candidate access paths, both checked here since neither is
confirmed working yet:
  1. data.stats.gov.cn's own query engine (the "National Data"
     database behind the government's public data portal) - an
     AJAX-style endpoint (easyquery.htm) the site's own JS calls to
     populate query results. Might return clean JSON directly, might
     require session/referer headers to avoid a bot-block, might not
     work at all without reverse-engineering internal indicator codes.
  2. The monthly "National Economy" statistical press release/bulletin
     pages (stats.gov.cn, English section) - narrative reports that
     include a numbers table for "Output of Major Industrial
     Products" (raw coal, crude oil, natural gas, electricity) each
     month. Static HTML, no query engine needed, but format/URL
     pattern across months is unconfirmed.

This script just probes both paths and reports what's actually
reachable/parseable - no assumptions carried in from outside this run.
"""

import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
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


def probe_data_portal():
    # The English "National Data" landing page - just to confirm the
    # domain/section is reachable at all, and see what real query-page
    # URLs/asset paths it references (its JS likely reveals the real
    # easyquery.htm parameter format).
    r = try_get("data.stats.gov.cn English portal", "https://data.stats.gov.cn/english/")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])

    # Monthly indicator query engine, guessed parameters based on the
    # site's known "hgyd" (monthly) dataset code - likely wrong first
    # try, but the exact error/response shape will show what's needed.
    r = try_get(
        "data.stats.gov.cn easyquery (monthly, guessed params)",
        "https://data.stats.gov.cn/english/easyquery.htm",
        params={"m": "QueryData", "dbcode": "hgyd", "rowcode": "zb", "colcode": "sj", "wds": "[]", "dfwds": "[]"},
    )
    if r is not None:
        print(r.text[:3000])


def probe_press_releases():
    # English press-release index - looking for the monthly economy
    # bulletin's real URL pattern.
    r = try_get("stats.gov.cn English press release index",
                "https://www.stats.gov.cn/english/PressRelease/")
    if r is not None and r.status_code == 200:
        print(r.text[:5000])


def main():
    probe_data_portal()
    probe_press_releases()


if __name__ == "__main__":
    main()
