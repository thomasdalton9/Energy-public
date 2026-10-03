"""One-off probe: National Gas (GB) Data Portal - NTS physical flows / demand by sector (LDZ, power stations, industrial,
interconnectors, storage) and supply by terminal - what the API offers and how far back it goes. Prints samples only."""
import json
import re
import sys
import requests

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data-probe)"


def get(url, **kw):
    try:
        r = S.get(url, timeout=90, **kw)
        print(f"\nGET {r.url}\n -> {r.status_code} {len(r.content)} bytes {r.headers.get('content-type')}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"\nGET {url}\n -> ERROR {type(e).__name__}: {e}")


for u in ("https://data.nationalgas.com/", "https://data.nationalgas.com/find-gas-data",
          "https://data.nationalgas.com/api/gas-data-items", "https://data.nationalgas.com/api/get-data-items",
          "https://data.nationalgas.com/api/find-gas-data-download?applicableFor=Y&dateFrom=2026-09-20&dateTo=2026-09-28&dateType=GASDAY&latestFlag=Y&ids=PUBOB637&type=CSV",
          "https://data.nationalgas.com/api/find-gas-data-download?applicableFor=Y&dateFrom=2026-09-20&dateTo=2026-09-28&dateType=GASDAY&latestFlag=Y&ids=PUBOB637,PUBOBJ1&type=json"):
    r = get(u)
    if r is not None and r.status_code == 200:
        print(r.text[:1500])
        for m in sorted(set(re.findall(r'(?:src|href)="([^"]+\.js)"', r.text)))[:10]:
            print("  script:", m)
r = get("https://data.nationalgas.com/")
if r is not None and r.status_code == 200:
    for m in sorted(set(re.findall(r'(?:src|href)="([^"]+\.js)"', r.text)))[:6]:
        js = get(m if m.startswith("http") else "https://data.nationalgas.com" + m)
        if js is not None and js.status_code == 200:
            for pat in (r'api/[A-Za-z0-9_\-/]+', r'PUBOB[A-Z0-9]+'):
                found = sorted(set(re.findall(pat, js.text)))
                print(" ", pat, found[:60])
sys.exit(0)
