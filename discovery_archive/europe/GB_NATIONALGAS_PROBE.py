"""One-off probe 2: National Gas (GB) Data Portal item tree - ids for NTS supply by terminal / LNG / storage, demand by
sector (LDZ, power stations, industrial, interconnectors), and history depth."""
import json
import re
import sys
import requests

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data-probe)"
B = "https://data.nationalgas.com/api"


def get(url, **kw):
    try:
        r = S.get(url, timeout=120, **kw)
        print(f"\nGET {r.url}\n -> {r.status_code} {len(r.content)} bytes {r.headers.get('content-type')}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"\nGET {url}\n -> ERROR {type(e).__name__}: {e}")


for ep in ("find-gas-data-folders", "gas-data-reports-folders", "search-everywhere?q=power"):
    r = get(f"{B}/{ep}")
    if r is not None and r.status_code == 200:
        txt = r.text
        print(txt[:3000])
        try:
            j = r.json()
        except Exception:  # noqa: BLE001
            continue
        # walk the tree and print every node that has an id-like field
        out = []

        def walk(x, path=""):
            if isinstance(x, dict):
                name = x.get("name") or x.get("label") or x.get("title") or x.get("dataItem") or ""
                ident = x.get("id") or x.get("itemId") or x.get("publicationObjectName") or x.get("publicationObjectId")
                if ident and name:
                    out.append((path + "/" + str(name), ident))
                for k, v in x.items():
                    walk(v, path + "/" + str(name) if name else path)
            elif isinstance(x, list):
                for v in x:
                    walk(v, path)
        walk(j)
        print(f"-- {len(out)} named ids; those matching demand/supply/power/LDZ/terminal/storage/LNG/interconnector:")
        for p, i in out:
            if re.search(r"demand|supply|power|ldz|terminal|storage|lng|interconnector|industrial|shrink|beach", p, re.I):
                print("  ", i, p)
for ids in ("PUBOB637", "PUBOB39", "PUBOB42"):
    r = get(f"{B}/find-gas-data-download?applicableFor=Y&dateFrom=2021-01-01&dateTo=2021-01-05&dateType=GASDAY&latestFlag=Y&ids={ids}&type=CSV")
    if r is not None and r.status_code == 200:
        print(r.text[:500])
sys.exit(0)
