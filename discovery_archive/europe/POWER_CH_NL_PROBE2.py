"""
Probe 2: Swiss (BFE energiedashboard / Swissgrid) and Dutch (CBS) raw electricity production sources.
CH: opendata.swiss CKAN package_search for 'swissgrid' / 'energiedashboard', then each resource URL with browser headers; direct uvek-gis CSVs
with full browser headers. NL: CBS catalogue (no $select), then the first tables' columns.
Prints only.
"""
import io
import json
import sys

import pandas as pd
import requests

BR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "text/csv,application/json,text/html,*/*;q=0.8", "Accept-Language": "en-GB,en;q=0.9", "Referer": "https://www.bfe.admin.ch/"}


def show_csv(url):
    r = requests.get(url, headers=BR, timeout=90)
    print(" ", url, r.status_code, len(r.content), r.headers.get("content-type", "")[:40])
    if not r.ok or r.text.lstrip().startswith("<"):
        print("   ->", r.text[:160].replace("\n", " "))
        return None
    df = pd.read_csv(io.StringIO(r.text))
    print("   cols:", df.columns.tolist()[:14], "rows", len(df))
    print("   head:", df.head(2).to_dict("records"))
    return df


print("=== CH: CKAN search")
for q in ("swissgrid", "energiedashboard", "Stromproduktion Swissgrid"):
    try:
        r = requests.get("https://ckan.opendata.swiss/api/3/action/package_search", params={"q": q, "rows": 12}, headers=BR, timeout=60)
        print("query", q, r.status_code)
        for p in r.json().get("result", {}).get("results", []):
            t = p.get("title")
            title = t.get("de") or t.get("en") if isinstance(t, dict) else t
            print("  PKG", p.get("name"), "|", str(title)[:90])
            for x in p.get("resources", [])[:4]:
                print("     ", x.get("format"), x.get("download_url") or x.get("url"))
    except Exception as e:  # noqa: BLE001
        print("CKAN FAILED", q, type(e).__name__, str(e)[:120])
print("\n=== CH: direct CSVs")
for u in ("https://www.uvek-gis.admin.ch/BFE/ogd/103/ogd103_stromverbrauch_swissgrid_lv_und_endv.csv",
          "https://www.uvek-gis.admin.ch/BFE/ogd/104/ogd104_stromproduktion_swissgrid.csv"):
    try:
        df = show_csv(u)
    except Exception as e:  # noqa: BLE001
        print("  FAILED", u, type(e).__name__, str(e)[:120])

print("\n=== NL: CBS catalogue")
seen = 0
for q in ("lektriciteit", "onnestroom", "ernieuwbare"):
    try:
        r = requests.get("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$format": "json", "$filter": f"substringof('{q}',Title) eq true"},
                         headers=BR, timeout=60)
        print("query", q, r.status_code)
        for t in r.json().get("value", [])[:40]:
            print("  ", t.get("Identifier"), "|", str(t.get("Title"))[:85], "|", t.get("Period"), "|", t.get("Frequency"), "|", t.get("Status"))
    except Exception as e:  # noqa: BLE001
        print("CBS FAILED", q, type(e).__name__, str(e)[:160])
sys.exit(0)
