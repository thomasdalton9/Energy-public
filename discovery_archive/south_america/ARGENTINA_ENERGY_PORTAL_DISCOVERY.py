"""
Discovery pass for Argentina's official open-data ecosystem
(datos.gob.ar / datos.energia.gob.ar), looking for:
  - Historical electricity generation by source (ideally back to 2015,
    to extend argentina_generation_mix.py which is CAMMESA-sourced and
    hard-limited to Oct-2024 onwards for daily generator-level data) -
    a web search surfaced "Generacion Electrica Bruta Asociada a Redes
    Desde 1930" on datos.gob.ar, suggesting a long official series exists.
  - Natural gas production
  - Natural gas demand by category/sector (residential, industrial,
    power generation, CNG, exports...)
  - Electricity and/or gas exports

datos.gob.ar is CKAN-based (like data.kapsarc.org earlier this session,
though hopefully with a working search this time) - its public API needs
no auth: https://datos.gob.ar/api/3/action/package_search
Argentina's Ministry of Economy also runs a separate clean JSON time-series
API at https://apis.datos.gob.ar/series/api/ - worth checking directly
since it may have the same underlying series in a much easier format
than a CKAN CSV resource.

This script just probes both and reports what's actually there - no
assumptions carried in from outside this run.
"""

import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)

CKAN_SEARCH_URL = "https://datos.gob.ar/api/3/action/package_search"
SERIES_SEARCH_URL = "https://apis.datos.gob.ar/series/api/search/"
SERIES_DATA_URL = "https://apis.datos.gob.ar/series/api/series/"

QUERIES = [
    "generacion electrica por fuente",
    "generacion electrica bruta asociada a redes",
    "produccion de gas natural",
    "demanda de gas natural por prestacion",
    "consumo de gas natural por categoria",
    "exportacion de energia electrica",
    "exportacion de gas natural",
    "importacion de energia electrica",
    "intercambio internacional de electricidad",
]


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code}  bytes={len(r.content)}  content-type={r.headers.get('content-type')}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def search_ckan(query):
    r = try_get(f"datos.gob.ar CKAN search: {query!r}", CKAN_SEARCH_URL, params={"q": query, "rows": 8})
    if r is None or r.status_code != 200:
        return
    try:
        data = r.json()
    except ValueError:
        print("  response wasn't valid JSON")
        print(r.text[:500])
        return
    result = data.get("result", {})
    total = result.get("count", "?")
    packages = result.get("results", [])
    print(f"  total count={total}, {len(packages)} returned:")
    for pkg in packages:
        name = pkg.get("name")
        title = pkg.get("title")
        org = (pkg.get("organization") or {}).get("title")
        resources = pkg.get("resources", [])
        print(f"    [{name}] {title}  (org: {org}, {len(resources)} resources)")
        for res in resources[:6]:
            print(f"        - {res.get('format')}: {res.get('name')} -> {res.get('url')}")


def probe_ckan_query_actually_works():
    # Same bug class as KAPSARC earlier this session - confirm the q
    # param really filters before trusting any result above.
    search_ckan("zzznonexistentqueryzzz12345")


def search_series_api(query):
    r = try_get(f"Series de Tiempo API search: {query!r}", SERIES_SEARCH_URL, params={"q": query, "limit": 8})
    if r is None or r.status_code != 200:
        return
    try:
        data = r.json()
    except ValueError:
        print("  response wasn't valid JSON")
        print(r.text[:500])
        return
    results = data if isinstance(data, list) else data.get("data", data)
    print(f"  response type={type(results)}: {str(results)[:1500]}")


def probe_series_metadata_csv():
    # The Series de Tiempo API also publishes a full metadata catalog CSV
    # listing every series id, title, units, and date range - much more
    # reliable than trying to guess the free-text search endpoint's exact
    # shape.
    r = try_get(
        "Series de Tiempo full catalog CSV",
        "https://apis.datos.gob.ar/series/api/dump/series-tiempo-metadatos.csv",
    )
    if r is None or r.status_code != 200:
        return
    text = r.content.decode("utf-8", errors="replace")
    lines = text.splitlines()
    print(f"  {len(lines)} lines total")
    header = lines[0] if lines else ""
    print(f"  header: {header}")
    keywords = ["electric", "gas natural", "generacion", "exportacion", "importacion", "demanda de gas", "produccion de gas"]
    hits = [line for line in lines[1:] if any(k in line.lower() for k in keywords)]
    print(f"  {len(hits)} lines mentioning any of {keywords}:")
    for line in hits[:60]:
        print(f"    {line[:300]}")


def main():
    for q in QUERIES:
        search_ckan(q)
    probe_ckan_query_actually_works()
    for q in QUERIES[:4]:
        search_series_api(q)
    probe_series_metadata_csv()


if __name__ == "__main__":
    main()
