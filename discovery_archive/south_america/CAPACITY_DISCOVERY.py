"""
Check whether two already-confirmed South America APIs also expose
installed/nameplate capacity (MW) per plant, not just generation:

  Colombia (XM): COLOMBIA_XM_GENERATION.py already calls
  /lists MetricId=ListadoRecursos, Entity=Sistema for each plant's fuel
  type - this just dumps every column that response actually has, to
  see if a capacity field (e.g. "Capacidad Efectiva Neta"/MPO) is there.

  Brazil (ANEEL): ANEEL's SIGA (Sistema de Informacoes de Geracao) is a
  well-known public CKAN dataset listing every generation unit in
  Brazil with installed capacity, fuel, and location. Probing its CKAN
  API root (dadosabertos.aneel.gov.br) since this project hasn't called
  it before - confirming reachability and real resource/field names
  before writing anything that depends on it.

Usage: python3 CAPACITY_DISCOVERY.py
"""
import sys

import requests

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url} {kwargs.get('params', '')}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def colombia_xm_listado_recursos():
    print(f"\n{'=' * 70}\nXM ListadoRecursos (POST): https://servapibi.xm.com.co/lists\n{'=' * 70}", file=sys.stderr)
    r = requests.post("https://servapibi.xm.com.co/lists", headers=HEADERS, timeout=TIMEOUT,
                       json={"MetricId": "ListadoRecursos", "Entity": "Sistema"})
    print(f"  POST status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
    if r.status_code != 200:
        print(r.text[:500], file=sys.stderr)
        return
    data = r.json()
    print(f"  top-level keys: {list(data.keys())}", file=sys.stderr)
    items = data.get("Items") or data.get("items") or []
    if items:
        first = items[0]
        print(f"  first item keys: {list(first.keys())}", file=sys.stderr)
        # drill into the nested records structure if present
        for key, val in first.items():
            if isinstance(val, list) and val:
                print(f"  first['{key}'][0]: {val[0]}", file=sys.stderr)


def aneel_siga_probe():
    base = "https://dadosabertos.aneel.gov.br/api/3/action/"
    r = try_get("ANEEL CKAN package_search (siga)", base + "package_search", params={"q": "SIGA geracao", "rows": 10})
    if r is not None and r.status_code == 200:
        try:
            data = r.json()
            results = data.get("result", {}).get("results", [])
            print(f"  {len(results)} package(s) found", file=sys.stderr)
            for pkg in results[:5]:
                print(f"    {pkg.get('name')}: {pkg.get('title')}", file=sys.stderr)
                for res in pkg.get("resources", [])[:5]:
                    print(f"      resource: {res.get('name')} format={res.get('format')} url={res.get('url')}",
                          file=sys.stderr)
        except ValueError:
            print(r.text[:1000], file=sys.stderr)


if __name__ == "__main__":
    print("=== Colombia XM ListadoRecursos ===", file=sys.stderr)
    try:
        colombia_xm_listado_recursos()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Brazil ANEEL SIGA ===", file=sys.stderr)
    try:
        aneel_siga_probe()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
