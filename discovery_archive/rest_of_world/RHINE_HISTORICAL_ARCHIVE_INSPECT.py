"""
Follow-up to RHINE_HISTORY_DEPTH_CHECK.py: PEGELONLINE's regular
/measurements.json endpoint only supports up to a 30-day window
(start=P30D; P1Y/P5Y/P30Y all 500). A web search surfaced a separate
historical-archive endpoint:

    POST https://www.pegelonline.wsv.de/gast/historische-zeitreihen/prepare-download
    params: uuid, parameter, start (ISO-UTC), end (ISO-UTC), format (csv|json)
    -> a 303 redirect to a ZIP, daily resolution back to 2000-01-01

Testing it for Kaub's confirmed station UUID
(1d26e504-7f9e-480a-b52c-5932be6549ab), parameter W (water level), a
multi-year historical range.
"""
import sys

import requests

URL = "https://www.pegelonline.wsv.de/gast/historische-zeitreihen/prepare-download"
KAUB_UUID = "1d26e504-7f9e-480a-b52c-5932be6549ab"
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 45)

params = {
    "uuid": KAUB_UUID,
    "parameter": "W",
    "start": "2015-01-01T00:00:00.000Z",
    "end": "2020-12-31T23:59:59.000Z",
    "format": "csv",
}

print(f"POST {URL}\nparams={params}", file=sys.stderr)
r = requests.post(URL, headers=HEADERS, data=params, timeout=TIMEOUT, allow_redirects=False)
print(f"status={r.status_code}", file=sys.stderr)
print(f"headers: {dict(r.headers)}", file=sys.stderr)
if r.status_code in (301, 302, 303, 307, 308):
    redirect_url = r.headers.get("Location")
    print(f"redirects to: {redirect_url}", file=sys.stderr)
    if redirect_url:
        if redirect_url.startswith("/"):
            redirect_url = "https://www.pegelonline.wsv.de" + redirect_url
            print(f"resolved to absolute: {redirect_url}", file=sys.stderr)
        r2 = requests.get(redirect_url, headers=HEADERS, timeout=TIMEOUT)
        print(f"  follow-up status={r2.status_code} bytes={len(r2.content)} "
              f"content-type={r2.headers.get('content-type')}", file=sys.stderr)
        print(f"  first 300 bytes: {r2.content[:300]}", file=sys.stderr)
else:
    print(f"body (first 500 bytes): {r.content[:500]}", file=sys.stderr)
