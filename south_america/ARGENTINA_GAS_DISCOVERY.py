"""
Follow-up on the national gas production/consumption data found earlier
this session via Argentina's Series de Tiempo API
(apis.datos.gob.ar/series/api/series/) - series 364.3_PRODUCCIoNRAL__25
(monthly gas production, 1996-2026) and 364.1_PRODUCCIONRAL__25 (annual).
A CSV named actividad-gas.csv was seen with many more columns
(consumption by category: residencial/comercial/industria/centrales_
electricas/... and by distributor: metrogas/camuzzi_gas/gasnor/...) but
this script never recorded where that CSV actually comes from - rather
than guess a download URL, this queries the series API's own metadata
and search endpoints (both already confirmed working) to find the real
series ids/distribution links for those other columns.

Usage: python3 ARGENTINA_GAS_DISCOVERY.py
"""
import sys

import requests

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)
SERIES_API = "https://apis.datos.gob.ar/series/api/series/"
SEARCH_API = "https://apis.datos.gob.ar/series/api/search/"

KNOWN_IDS = ["364.3_PRODUCCIoNRAL__25", "364.1_PRODUCCIONRAL__25"]


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url} {kwargs.get('params')}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def metadata_for_known_ids():
    r = try_get("full metadata for known series ids", SERIES_API,
                params={"ids": ",".join(KNOWN_IDS), "metadata": "full", "limit": 0})
    if r is not None and r.status_code == 200:
        try:
            data = r.json()
            print(f"  {str(data)[:4000]}", file=sys.stderr)
        except ValueError:
            print(r.text[:1000], file=sys.stderr)


def search_gas_series():
    for q in ["gas natural", "consumo de gas", "produccion de gas natural", "actividad-gas"]:
        r = try_get(f"search: {q!r}", SEARCH_API, params={"q": q, "limit": 20})
        if r is not None and r.status_code == 200:
            try:
                print(f"  {str(r.json())[:3000]}", file=sys.stderr)
            except ValueError:
                print(r.text[:500], file=sys.stderr)


def main():
    metadata_for_known_ids()
    search_gas_series()


if __name__ == "__main__":
    main()
