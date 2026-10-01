"""
Why does the GNI daily gas demand archive stop at 2026-03-31 when the
resolved file's name embeds 2026-08? Dumps EVERY resource in the
data.gov.ie packages for gas demand and gas supply (name, URL, format,
created/last_modified), fetches each CSV and prints its date range, and
probes GNI's open-data folder directly for newer quarterly files - to
find where the post-March data lives and fix the resolvers to pick it.
"""
import io
import sys

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 60)
CKAN = "https://data.gov.ie/api/3/action/package_show"


def date_range(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            return f"HTTP {r.status_code}"
        df = pd.read_csv(io.BytesIO(r.content))
        col = df.columns[0]
        d = pd.to_datetime(df[col], dayfirst=True, errors="coerce").dropna()
        return f"{len(df)} rows, {d.min().date()} to {d.max().date()}, cols={list(df.columns)}"
    except Exception as e:
        return f"ERR {type(e).__name__}: {e}"


for pkg in ["dailygasdemandireland", "dailygassupply"]:
    print(f"\n=== CKAN package {pkg} ===", flush=True)
    r = requests.get(CKAN, headers=HEADERS, params={"id": pkg}, timeout=TIMEOUT)
    result = r.json()["result"]
    print(f"metadata_modified={result.get('metadata_modified')} num_resources={result.get('num_resources')}", flush=True)
    for res in result["resources"]:
        print(f"- {res.get('name')!r} format={res.get('format')} created={res.get('created')} "
              f"last_modified={res.get('last_modified')}\n    url={res.get('url')}", flush=True)
        if (res.get("format") or "").upper() == "CSV":
            print(f"    -> {date_range(res['url'])}", flush=True)

print("\n=== Probing GNI open-data folder for quarterly files ===", flush=True)
for kind in ["Demand", "Supply"]:
    for y in [2025, 2026]:
        for q in [1, 2, 3, 4]:
            for pattern in [f"https://www.gasnetworks.ie/corporate/open-data/{y}-Q{q}-Daily-Gas-{kind}.csv",
                            f"https://www.gasnetworks.ie/corporate/open-data/Daily-Gas-{kind}-{y}-Q{q}.csv"]:
                try:
                    h = requests.head(pattern, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
                    if h.status_code == 200:
                        print(f"FOUND {pattern} -> {date_range(pattern)}", flush=True)
                    else:
                        print(f"  {h.status_code} {pattern}", flush=True)
                except Exception as e:
                    print(f"  ERR {pattern}: {type(e).__name__}", flush=True)

print("\n=== GNI open-data page links ===", flush=True)
for page in ["https://www.gasnetworks.ie/corporate/open-data/", "https://www.gasnetworks.ie/corporate/gas-regulation/transparency/"]:
    try:
        r = requests.get(page, headers=HEADERS, timeout=TIMEOUT)
        import re
        links = sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx?))"', r.text, re.I)))
        print(f"{page}: HTTP {r.status_code}, {len(links)} csv/xlsx links", flush=True)
        for l in links:
            print("   ", l, flush=True)
    except Exception as e:
        print(f"{page}: ERR {type(e).__name__}: {e}", flush=True)
