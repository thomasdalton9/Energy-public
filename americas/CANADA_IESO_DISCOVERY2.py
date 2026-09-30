"""
Follow-up to CANADA_IESO_DISCOVERY.py: reports.ieso.ca returned a
plain 404 (that domain is dead). A web search found the current public
reports domain is reports-public.ieso.ca, with report names unchanged
(e.g. GenOutputbyFuelHourly, per its own helpfile PDF at
reports-public.ieso.ca/docrefs/helpfile/GenOutputbyFuelMonthly_h1.pdf).
Checking that domain's real report directories live.
"""
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
        if r.status_code == 200:
            print(f"  first 1000 chars: {r.text[:1000]}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


if __name__ == "__main__":
    try_get("GenOutputbyFuelHourly directory listing", "https://reports-public.ieso.ca/public/GenOutputbyFuelHourly/")
    try_get("GenOutputCapability directory listing", "https://reports-public.ieso.ca/public/GenOutputCapability/")
    try_get("PUB list root", "https://reports-public.ieso.ca/public/")
