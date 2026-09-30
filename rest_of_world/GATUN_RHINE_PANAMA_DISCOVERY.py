"""
Discovery pass for three leads found via web search (none confirmed
live yet):

  Gatun Lake (Panama Canal): ACP's own "Gatun Water Level Indicators"
  dashboard (evtms-rpts.pancanal.com/eng/h2o/index.html) reportedly
  offers Historical Water Levels and Water Level Projections as CSV,
  with data back to 1965. Checking what's actually behind that page and
  whether a CSV download link is a real, fetchable URL.

  Rhine river (Kaub gauge): PEGELONLINE (pegelonline.wsv.de), the German
  federal waterways administration's REST/JSON API - free, no key,
  documented at pegelonline.wsv.de/webservice/dokuRestapi. Checking the
  real station list (to find Kaub's station UUID) and a sample reading.

  Panama power generation: CND (Centro Nacional de Despacho) operates a
  "SITR" (Sistema de Intercambio de Informacion en Tiempo Real) at
  sitr.cnd.com.pa - checking whether it's a plain page or a JS app like
  Bolivia/Sri Lanka were (in which case Playwright network capture would
  be the next step, not plain requests).

Usage: python3 GATUN_RHINE_PANAMA_DISCOVERY.py
"""
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


def gatun_lake():
    r = try_get("ACP Gatun water level dashboard", "https://evtms-rpts.pancanal.com/eng/h2o/index.html")
    if r is not None and r.status_code == 200:
        text = r.text
        import re
        links = sorted(set(re.findall(r'href="([^"]+\.csv[^"]*)"', text, re.I)))
        print(f"  CSV links found in page: {links}", file=sys.stderr)
        if not links:
            print(f"  first 1500 chars:\n{text[:1500]}", file=sys.stderr)
        for link in links[:3]:
            full_url = link if link.startswith("http") else f"https://evtms-rpts.pancanal.com{link}"
            r2 = try_get("candidate CSV", full_url)
            if r2 is not None and r2.status_code == 200:
                print(f"  first 500 bytes: {r2.content[:500]}", file=sys.stderr)


def rhine_pegelonline():
    r = try_get("PEGELONLINE station list", "https://www.pegelonline.wsv.de/webservices/rest-api/v2/stations.json",
                params={"waters": "RHEIN"})
    if r is not None and r.status_code == 200:
        try:
            stations = r.json()
            print(f"  {len(stations)} Rhine stations", file=sys.stderr)
            kaub = [s for s in stations if "kaub" in str(s.get("longname", "")).lower()]
            print(f"  Kaub match: {kaub}", file=sys.stderr)
            if kaub:
                uuid = kaub[0]["uuid"]
                r2 = try_get("Kaub current water level (W)",
                              f"https://www.pegelonline.wsv.de/webservices/rest-api/v2/stations/{uuid}/W/measurements.json",
                              params={"start": "P1D"})
                if r2 is not None and r2.status_code == 200:
                    print(f"  sample: {r2.json()[:5]}", file=sys.stderr)
        except ValueError:
            print(r.text[:1000], file=sys.stderr)


def panama_cnd_sitr():
    r = try_get("Panama CND SITR", "https://sitr.cnd.com.pa/")
    if r is not None and r.status_code == 200:
        text = r.text
        print(f"  page length: {len(text)} chars", file=sys.stderr)
        print(f"  first 1500 chars:\n{text[:1500]}", file=sys.stderr)


def puerto_rico_eia():
    """americas/EIA930_FUEL_MIX_DAILY.py already confirms api.eia.gov/v2/...
    works with EIA_API_KEY read from the environment (a GH Actions
    secret in this repo). electric-power-operational-data is EIA's
    documented state-level monthly generation-by-fuel category
    (EIA-923-based) - Puerto Rico is tracked there as location "PR".
    Confirming that live rather than assuming it."""
    import os
    api_key = os.environ.get("EIA_API_KEY")
    if not api_key:
        print("  EIA_API_KEY not set in environment - skipping", file=sys.stderr)
        return
    r = try_get(
        "EIA electric-power-operational-data for PR",
        "https://api.eia.gov/v2/electricity/electric-power-operational-data/data/",
        params={
            "api_key": api_key,
            "frequency": "monthly",
            "data[]": "generation",
            "facets[location][]": "PR",
            "sort[0][column]": "period",
            "sort[0][direction]": "desc",
            "length": 20,
        },
    )
    if r is not None and r.status_code == 200:
        try:
            data = r.json()
            rows = data.get("response", {}).get("data", [])
            print(f"  {len(rows)} rows returned", file=sys.stderr)
            for row in rows[:10]:
                print(f"    {row}", file=sys.stderr)
        except ValueError:
            print(r.text[:1000], file=sys.stderr)
    elif r is not None:
        print(r.text[:1000], file=sys.stderr)


if __name__ == "__main__":
    print("=== Gatun Lake (Panama Canal) ===", file=sys.stderr)
    try:
        gatun_lake()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Rhine river (PEGELONLINE) ===", file=sys.stderr)
    try:
        rhine_pegelonline()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Panama CND SITR ===", file=sys.stderr)
    try:
        panama_cnd_sitr()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Puerto Rico via EIA ===", file=sys.stderr)
    try:
        puerto_rico_eia()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
