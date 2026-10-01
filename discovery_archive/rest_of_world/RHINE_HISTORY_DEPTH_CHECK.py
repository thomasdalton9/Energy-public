"""
PEGELONLINE (the German federal waterways REST API confirmed working
for the Rhine's Kaub gauge) is primarily a real-time service - checking
how far back its own /measurements.json endpoint actually goes before
assuming multi-year history is available directly from it. If it's
short (common for these "current conditions" APIs), a separate
long-history source will be needed for a 5-year-range chart.
"""
import sys

import requests

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)
KAUB_UUID = "1d26e504-7f9e-480a-b52c-5932be6549ab"


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url} {kwargs.get('params', '')}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def main():
    base = f"https://www.pegelonline.wsv.de/webservices/rest-api/v2/stations/{KAUB_UUID}/W/measurements.json"
    for start in ["P30D", "P1Y", "P5Y", "P30Y"]:
        r = try_get(f"Kaub W measurements, start={start}", base, params={"start": start})
        if r is not None and r.status_code == 200:
            try:
                data = r.json()
                print(f"  {len(data)} readings. First: {data[0] if data else None}  Last: {data[-1] if data else None}",
                      file=sys.stderr)
            except ValueError:
                print(r.text[:500], file=sys.stderr)

    # also check the station detail endpoint for any metadata about
    # available history / gauge zero reference
    r = try_get("Kaub station detail", f"https://www.pegelonline.wsv.de/webservices/rest-api/v2/stations/{KAUB_UUID}.json")
    if r is not None and r.status_code == 200:
        print(f"  {r.json()}", file=sys.stderr)


if __name__ == "__main__":
    main()
