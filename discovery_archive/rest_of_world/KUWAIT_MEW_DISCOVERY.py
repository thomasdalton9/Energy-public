"""
Discovery pass for Kuwait electricity data via MEW (Ministry of
Electricity, Water and Renewable Energy) - a SEPARATE, dedicated
parser from electricitymaps-contrib's KW.py, distinct from the
already-confirmed-dead GCCIA.py (shared Gulf-wide consumption scraper,
site redesigned, live widget gone).

KW.py's own approach: fetch https://www.mew.gov.kw/en directly and
regex-search for a 4-5 digit number in parentheses - "(XXXXX)" - which
is the current system load (MW) shown on a live gauge widget on that
page. No generation-by-fuel breakdown (consumption is treated as a
proxy for production); the parser's own comment cites a static IEA
2017 split (65.6% oil, 34.4% gas) as background, not live data.

This script just probes whether the page and pattern still work - no
assumptions carried in from outside this run.
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


def probe_mew_load_gauge():
    r = try_get("MEW English homepage", "https://www.mew.gov.kw/en")
    if r is None or r.status_code != 200:
        return
    matches = re.findall(r"\((\d{4,5})\)", r.text)
    print(f"  {len(matches)} 4-5-digit-in-parens matches found: {matches}")
    if not matches:
        print("  no matches - dumping full page for manual inspection")
        print(r.text)
    else:
        # show a bit of context around each match to sanity-check it's
        # really the load gauge and not some unrelated number
        for m in re.finditer(r".{80}\(\d{4,5}\).{20}", r.text):
            print(f"  context: ...{m.group(0)}...")


def probe_mew_navigation():
    # in case the load gauge moved to a different page/widget since
    # KW.py was last verified (same pattern as Cyprus's site update) -
    # scanning the homepage's own links for anything load/generation/
    # consumption-related as a fallback lead
    r = try_get("MEW homepage (for nav scan)", "https://www.mew.gov.kw/en")
    if r is None or r.status_code != 200:
        return
    links = sorted(set(re.findall(r'href=["\']([^"\']*(?:load|generation|consumption|statistic|data)[^"\']*)["\']', r.text, re.I)))
    print(f"  {len(links)} nav links mentioning load/generation/consumption/statistic/data:")
    for l in links[:30]:
        print(f"    {l}")


def main():
    probe_mew_load_gauge()
    probe_mew_navigation()


if __name__ == "__main__":
    main()
