"""
Follow-up to IRELAND_TURKEY_POWER_PLAYWRIGHT_DISCOVERY.py: EPIAS's
homepage (seffaflik.epias.com.tr) shows a "Gerçek Zamanlı Üretim"
(Real-Time Generation) widget without requiring login, backed by a GET
(not the authenticated POST bulk endpoint turkey_generation_mix.py
already uses) to:

    https://seffaflik.epias.com.tr/electricity-service/v1/dashboard/realtime-generation

Checking what this actually returns - if it's a real, unauthenticated
current generation-by-source snapshot, it could give a no-credentials
"latest" reading even while turkey_generation_mix.py's historical
backfill stays blocked on a free EPIAS account's password.
"""
import json
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://seffaflik.epias.com.tr/",
}
TIMEOUT = (10, 45)

URL = "https://seffaflik.epias.com.tr/electricity-service/v1/dashboard/realtime-generation"

r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
print(f"status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
      file=sys.stderr)
if r.status_code == 200:
    try:
        data = r.json()
        print(json.dumps(data, indent=2)[:3000], file=sys.stderr)
    except ValueError:
        print(r.text[:1000], file=sys.stderr)
else:
    print(r.text[:1000], file=sys.stderr)
