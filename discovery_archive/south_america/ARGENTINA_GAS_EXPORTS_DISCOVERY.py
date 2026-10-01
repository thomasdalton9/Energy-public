"""
Discovery for Argentina natural gas EXPORTS by destination (Chile, Brazil,
Uruguay, Bolivia, ...), monthly from 2021, to add to argentina_gas_monthly.xlsx.

Probes (prints only compact facts so the job log isn't truncated):
  A. Series de Tiempo API search (apis.datos.gob.ar/series/api/search/)
  B. CKAN package_search on datos.gob.ar and datos.energia.gob.ar
  C. ENARGAS "Datos operativos" pages - links mentioning export / xls
  D. Download candidate xlsx/csv files found in B/C and print their layout

Usage: python3 ARGENTINA_GAS_EXPORTS_DISCOVERY.py
"""
import io
import re
import sys
from urllib.parse import urljoin

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 60)
SERIES_SEARCH = "https://apis.datos.gob.ar/series/api/search/"

SERIES_QUERIES = ["exportacion gas natural", "exportaciones gas natural", "gas exportado", "exportacion gas chile",
                  "exportacion gas brasil", "exportacion gas uruguay", "comercio exterior gas",
                  "gas natural exportaciones", "importacion gas natural"]
CKAN_BASES = ["https://datos.gob.ar", "http://datos.energia.gob.ar"]
CKAN_QUERIES = ["exportacion gas", "exportaciones gas natural", "comercio exterior gas", "gas exportado",
                "exportaciones", "gas natural"]
ENARGAS_PAGES = [
    "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/datos-operativos.php",
    "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/datos-operativos-subsecciones.php",
] + [f"https://www.enargas.gob.ar/secciones/transporte-y-distribucion/datos-operativos-subsecciones.php?sec={i}"
     for i in range(1, 13)] + [
    "https://www.enargas.gob.ar/secciones/datos-abiertos/datos-abiertos.php",
    "https://www.enargas.gob.ar/secciones/informacion-estadistica/informacion-estadistica.php",
]

candidates = []  # (label, url)


def get(url, **kw):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
        return r
    except requests.RequestException as e:
        print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:150]}")
        return None


def series_search():
    print("\n===== A. Series de Tiempo search =====")
    seen = set()
    for q in SERIES_QUERIES:
        r = get(SERIES_SEARCH, params={"q": q, "limit": 50})
        if r is None or r.status_code != 200:
            print(f"q={q!r}: status {getattr(r, 'status_code', None)}")
            continue
        hits = r.json().get("data", [])
        print(f"q={q!r}: {len(hits)} hits")
        for h in hits:
            f, d = h.get("field", {}), h.get("dataset", {})
            if f.get("id") in seen:
                continue
            seen.add(f.get("id"))
            text = (f.get("title", "") + " " + f.get("description", "") + " " + d.get("title", "")).lower()
            if "gas" not in text:
                continue
            print(f"  {f.get('id')} | {f.get('title')} | {f.get('description', '')[:80]} | units={f.get('units')} | "
                  f"{f.get('frequency')} {f.get('time_index_start')}..{f.get('time_index_end')} | "
                  f"ds={d.get('title', '')[:60]} | pub={d.get('publisher', {}).get('name', '') if isinstance(d.get('publisher'), dict) else d.get('source', '')}")


def ckan_search():
    print("\n===== B. CKAN package_search =====")
    seen = set()
    for base in CKAN_BASES:
        for q in CKAN_QUERIES:
            r = get(f"{base}/api/3/action/package_search", params={"q": q, "rows": 30})
            if r is None or r.status_code != 200:
                print(f"{base} q={q!r}: status {getattr(r, 'status_code', None)}")
                continue
            try:
                res = r.json()["result"]["results"]
            except Exception as e:
                print(f"{base} q={q!r}: bad json {e}")
                continue
            print(f"{base} q={q!r}: {len(res)} packages")
            for p in res:
                key = (base, p.get("name"))
                if key in seen:
                    continue
                seen.add(key)
                title = p.get("title", "")
                blob = (title + " " + p.get("notes", "")).lower()
                if "gas" not in blob and "export" not in blob:
                    continue
                print(f"  PKG {p.get('name')} | {title[:90]} | org={(p.get('organization') or {}).get('title', '')[:40]}")
                for res_ in p.get("resources", []):
                    name, url = res_.get("name", ""), res_.get("url", "")
                    print(f"     - {name[:70]} [{res_.get('format')}] {url}")
                    if "export" in (name + url + title).lower() and "gas" in (name + url + title).lower():
                        candidates.append((f"ckan:{name[:40]}", url))


def enargas():
    print("\n===== C. ENARGAS pages =====")
    for page in ENARGAS_PAGES:
        r = get(page)
        if r is None:
            continue
        print(f"{page}: status={r.status_code} bytes={len(r.content)}")
        if r.status_code != 200:
            continue
        html = r.text
        title = re.search(r"<title>(.*?)</title>", html, re.S | re.I)
        print(f"  title: {title.group(1).strip()[:100] if title else ''}")
        for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, re.S | re.I):
            href, text = m.group(1), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            full = urljoin(page, href)
            low = (href + " " + text).lower()
            if any(k in low for k in ("export", ".xls", ".csv", "datos-operativos", "operativ")):
                print(f"   link: {text[:70]!r} -> {full}")
                if "export" in low and any(full.lower().endswith(x) for x in (".xls", ".xlsx", ".csv", ".zip")):
                    candidates.append((f"enargas:{text[:40]}", full))


def inspect_files():
    print("\n===== D. Candidate files =====")
    import pandas as pd
    done = set()
    for label, url in candidates[:12]:
        if url in done:
            continue
        done.add(url)
        r = get(url)
        if r is None:
            continue
        print(f"\n--- {label} {url}: status={r.status_code} bytes={len(r.content)} ct={r.headers.get('content-type')}")
        if r.status_code != 200:
            continue
        try:
            if url.lower().endswith(".csv") or "csv" in (r.headers.get("content-type") or ""):
                df = pd.read_csv(io.BytesIO(r.content), sep=None, engine="python", encoding_errors="replace")
                print(f"  cols={list(df.columns)[:30]} rows={len(df)}")
                print(df.head(5).to_string()[:1500])
                print(df.tail(5).to_string()[:1500])
            else:
                xl = pd.ExcelFile(io.BytesIO(r.content))
                print(f"  sheets={xl.sheet_names[:20]}")
                for sh in xl.sheet_names[:3]:
                    df = xl.parse(sh, header=None)
                    print(f"  [{sh}] shape={df.shape}")
                    print(df.head(15).to_string(max_colwidth=25)[:2500])
        except Exception as e:
            print(f"  parse error: {type(e).__name__}: {str(e)[:200]}")


def main():
    series_search()
    ckan_search()
    enargas()
    inspect_files()
    print("\nDONE", file=sys.stderr)


if __name__ == "__main__":
    main()
