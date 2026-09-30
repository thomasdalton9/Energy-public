"""
Discovery pass for Kuwait, UAE, Oman, and Qatar electricity generation
data via KAPSARC (King Abdullah Petroleum Studies and Research Center,
data.kapsarc.org) - a Saudi-based MENA energy data hub that already
surfaced Kuwait/Oman water-and-electricity datasets in a web search,
after no national TSO/utility open-data portal could be found directly
for any of these four countries.

data.kapsarc.org's URL pattern (.../explore/dataset/<id>/?flg=en-gb)
is the standard Opendatasoft platform - which has a well-documented
public catalog search API (no auth needed for public datasets):
  https://data.kapsarc.org/api/v2/catalog/datasets?where=...
This lets us search the ENTIRE catalog for anything relevant to these
4 countries' electricity generation, rather than guessing dataset IDs
one at a time.

This script just probes the catalog and reports what's actually
there - no assumptions carried in from outside this run.
"""

import json
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)

CATALOG_API = "https://data.kapsarc.org/api/v2/catalog/datasets"

COUNTRIES = ["Kuwait", "UAE", "United Arab Emirates", "Oman", "Qatar", "Morocco"]


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code}  bytes={len(r.content)}  content-type={r.headers.get('content-type')}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def search_catalog(query):
    r = try_get(
        f"KAPSARC catalog search: {query!r}",
        CATALOG_API,
        params={"q": query, "rows": 20},
    )
    if r is None or r.status_code != 200:
        return
    try:
        data = r.json()
    except ValueError:
        print("  response wasn't valid JSON")
        print(r.text[:1000])
        return
    total = data.get("total_count", "?")
    datasets = data.get("datasets", [])
    print(f"  total_count={total}, {len(datasets)} returned:")
    for ds in datasets:
        meta = ds.get("dataset", {}).get("metas", {}).get("default", {})
        dsid = ds.get("dataset", {}).get("dataset_id")
        title = meta.get("title")
        print(f"    [{dsid}] {title}")


def main():
    for country in COUNTRIES:
        search_catalog(f"{country} electricity generation")
    # also a broader sweep for "power generation" / "fuel mix" across
    # the whole catalog, in case country names aren't in the dataset
    # titles/descriptions where the search indexes them
    search_catalog("power generation mix")
    search_catalog("GCC electricity")


if __name__ == "__main__":
    main()
