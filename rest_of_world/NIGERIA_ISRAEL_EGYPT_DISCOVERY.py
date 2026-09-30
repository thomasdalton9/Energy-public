"""
First discovery pass for three more countries, from web search leads:

  Nigeria: the grid operator (NESO under TCN, recently reorganised as
  NISO - Nigerian Independent System Operator) has published daily
  generation operational reports on its own website since 2014 - but
  search results don't confirm a structured API, just "reports on the
  website". Checking niso.ng/neso.ng/nsong.org directly for a real,
  structured (not PDF-only) data source.

  Israel: Noga (the Israel Independent System Operator, took over
  system operation from IEC in 2021) - checking noga-iso.co.il directly
  for a generation-mix dashboard/API, since search didn't surface one.

  Egypt: Egypt's Open Data Portal (egypt-odp.portaljs.com) has a
  confirmed "Electricity Generation by Source" dataset page - checking
  it directly for a real CSV download link and its actual granularity
  (search suggests annual, by fuel type - useful but not daily/granular).

Usage: python3 NIGERIA_ISRAEL_EGYPT_DISCOVERY.py
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)


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


def nigeria():
    for url in ["https://niso.ng/", "https://nsong.org/", "https://www.nsong.org/",
                "https://www.niso.ng/", "https://www.neso.ng/"]:
        r = try_get("Nigeria system operator candidate", url, allow_redirects=True)
        if r is not None and r.status_code == 200:
            text = r.text
            print(f"  page length: {len(text)}", file=sys.stderr)
            links = sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx|json|pdf)[^"]*)"', text, re.I)))[:15]
            print(f"  data-looking links: {links}", file=sys.stderr)
            print(f"  first 1000 chars:\n{text[:1000]}", file=sys.stderr)


def israel():
    r = try_get("Noga ISO homepage", "https://www.noga-iso.co.il/")
    if r is not None and r.status_code == 200:
        text = r.text
        print(f"  page length: {len(text)}", file=sys.stderr)
        links = sorted(set(re.findall(r'href="([^"]+)"', text)))
        relevant = [l for l in links if any(k in l.lower() for k in
                    ["generat", "ייצור", "production", "mix", "data", "api", "transparen"])][:20]
        print(f"  relevant-looking links: {relevant}", file=sys.stderr)


def egypt():
    r = try_get("Egypt ODP electricity generation page",
                "https://egypt-odp.portaljs.com/@egypt-odp/electricity-generation-by-source")
    if r is not None and r.status_code == 200:
        text = r.text
        print(f"  page length: {len(text)}", file=sys.stderr)
        links = sorted(set(re.findall(r'href="([^"]+\.csv[^"]*)"', text, re.I)))
        print(f"  CSV links: {links}", file=sys.stderr)
        if not links:
            print(f"  first 1500 chars:\n{text[:1500]}", file=sys.stderr)
        for link in links[:2]:
            full = link if link.startswith("http") else f"https://egypt-odp.portaljs.com{link}"
            r2 = try_get("candidate CSV", full)
            if r2 is not None and r2.status_code == 200:
                print(f"  first 500 bytes: {r2.content[:500]}", file=sys.stderr)


if __name__ == "__main__":
    print("=== Nigeria ===", file=sys.stderr)
    try:
        nigeria()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Israel ===", file=sys.stderr)
    try:
        israel()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Egypt ===", file=sys.stderr)
    try:
        egypt()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
