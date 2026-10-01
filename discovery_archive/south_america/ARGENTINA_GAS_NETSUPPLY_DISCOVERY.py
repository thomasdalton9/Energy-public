"""
Discovery for Argentina gas NET SUPPLY: measured gas injected into the
transport system from domestic basins, plus imports (Escobar LNG FSRU,
Bahia Blanca LNG, pipeline imports from Bolivia / Chile), monthly from
2021 and daily if possible. Replaces the gross-production line on the
Argentina demand chart.

Probes (compact output so the job log isn't truncated):
  A. ENARGAS daily exp/imp reports, tipo_list imp_dentro / imp_fuera
     (same AJAX endpoint the exports pull uses): headers + monthly sums
  B. ENARGAS funciones.js: every PD_* function and the .php endpoints it calls
  C. ENARGAS transporte-y-distribucion / datos operativos pages: every
     dod-* page, datos-estadisticos file and xls/xlsx link
  D. Download candidate datos-estadisticos files (recibido / inyectado /
     entregado / combustible / GNL / importacion) and print their layout
  E. Series de Tiempo + datos.energia.gob.ar CKAN searches for injection,
     delivered-to-system and import series

Usage: python3 ARGENTINA_GAS_NETSUPPLY_DISCOVERY.py
"""
import io
import re
from urllib.parse import urljoin

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 120)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
LIST = BASE + "partes-diarios-exp-imp-consulta-listado.php"
PAGE = BASE + "dod-partes-exp-imp-consulta.php"
JS = BASE + "funciones/js/funciones.js"

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 200)

S = requests.Session()
S.headers.update(HEADERS)


def flat(s):
    return re.sub(r"\s+", " ", s)


def get(url, **kw):
    try:
        return S.get(url, timeout=TIMEOUT, **kw)
    except requests.RequestException as e:
        print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:150]}")
        return None


def parse(page):
    heads = [flat(re.sub(r"<[^>]+>", "", re.sub(r"<br\s*/?>", " | ", h))).strip()
             for h in re.findall(r"<th[^>]*>(.*?)</th>", page, re.S | re.I)]
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, re.S | re.I):
        cells = [flat(re.sub(r"<[^>]+>", " ", c)).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        if cells and re.match(r"\d{2}/\d{2}/\d{4}", cells[0]):
            rows.append(cells)
    return heads, rows


def probe_imports():
    print("===== A. ENARGAS daily import reports =====")
    for tipo in ("imp_dentro", "imp_fuera", "imp", "importacion"):
        r = get(PAGE, params={"tipo": tipo})
        if r is not None:
            title = re.findall(r"<h[1-4][^>]*>(.*?)</h[1-4]>", r.text, re.S | re.I)
            print(f"page tipo={tipo}: status={r.status_code} bytes={len(r.text)} "
                  f"headings={[flat(re.sub('<[^>]+>', '', t)).strip()[:80] for t in title][:6]}")
            print("   tipo values on page:", sorted(set(re.findall(r"tipo=([a-z_]+)", r.text))))
    windows = [("2021-01-01", "2021-12-30"), ("2022-01-01", "2022-12-30"), ("2023-01-01", "2023-12-30"),
               ("2024-01-01", "2024-12-30"), ("2025-01-01", "2025-12-30"), ("2026-01-01", "2026-09-30")]
    for tipo in ("imp_dentro", "imp_fuera"):
        for d0, d1 in windows:
            try:
                r = S.post(LIST, data={"fecha_desde": d0, "fecha_hasta": d1, "tipo_list": tipo}, timeout=TIMEOUT)
            except requests.RequestException as e:
                print(f"{tipo} {d0}: ERROR {e}")
                continue
            heads, rows = parse(r.text)
            print(f"\n== {tipo} {d0}..{d1}: status={r.status_code} bytes={len(r.text)} rows={len(rows)} heads={heads}")
            if not rows:
                print("   text:", flat(re.sub(r"<[^>]+>", " ", r.text))[:400])
                continue
            print("   first row:", rows[0], " last row:", rows[-1])
            df = pd.DataFrame([x[1:] for x in rows], index=pd.to_datetime([x[0] for x in rows], format="%d/%m/%Y"),
                              columns=heads[1:len(rows[0])] if len(heads) >= len(rows[0]) else None)
            df = df.apply(lambda c: pd.to_numeric(c.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                                                  errors="coerce"))
            m = df.resample("MS").sum() / 1000
            print("   monthly sums (million m3):")
            print(m.round(1).to_string())


def probe_js():
    print("\n===== B. funciones.js =====")
    r = get(JS)
    if r is None:
        return
    js = r.text
    print(f"bytes={len(js)}")
    print("functions:", re.findall(r"function\s+(\w+)", js))
    print("php endpoints:", sorted(set(re.findall(r"[\w\-/]+\.php", js))))


LINK_KEYS = ("dod-", "datos-estadisticos", ".xls", ".csv", "operativ", "recib", "inyec", "entreg", "combust",
             "gnl", "import", "cuenca", "parte")


def probe_pages():
    print("\n===== C. ENARGAS pages =====")
    pages = [BASE + "datos-operativos.php", BASE + "datos-operativos-subsecciones.php",
             BASE + "transporte-y-distribucion.php", BASE + "dod.php", BASE + "datos-operativos-diarios.php",
             "https://www.enargas.gob.ar/secciones/informacion-tecnica/datos-operativos.php",
             "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/dod-gas-recibido.php",
             "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/dod-partes-diarios.php",
             "https://www.enargas.gob.ar/secciones/datos-abiertos/datos-abiertos.php",
             "https://www.enargas.gob.ar/secciones/informacion-estadistica/informacion-estadistica.php",
             "https://www.enargas.gob.ar/"] + \
            [BASE + f"datos-operativos-subsecciones.php?sec={i}" for i in range(1, 21)]
    seen_pages, links = set(), {}
    queue = list(pages)
    while queue and len(seen_pages) < 70:
        u = queue.pop(0)
        if u in seen_pages:
            continue
        seen_pages.add(u)
        r = get(u)
        if r is None:
            continue
        h1 = re.findall(r"<h[1-3][^>]*>(.*?)</h[1-3]>", r.text, re.S | re.I)
        print(f"{r.status_code} {len(r.text):>7} {u}  {[flat(re.sub('<[^>]+>', '', t)).strip()[:60] for t in h1][:4]}")
        if r.status_code != 200:
            continue
        for m in re.finditer(r"<a[^>]+href=[\"']([^\"'#]+)[\"'][^>]*>(.*?)</a>", r.text, re.I | re.S):
            href, text = urljoin(u, m.group(1).strip()), flat(re.sub(r"<[^>]+>", " ", m.group(2))).strip()
            if "enargas.gob.ar" not in href:
                continue
            low = (href + " " + text).lower()
            if any(k in low for k in LINK_KEYS):
                links.setdefault(href, text[:70])
                if href.endswith(".php") or ".php?" in href:
                    if ("dod-" in href or "datos-operativos" in href) and href not in seen_pages:
                        queue.append(href)
    print(f"\n{len(links)} links:")
    for h, t in sorted(links.items()):
        print(f"  {t!r:72} {h}")
    return links


FILE_KEYS = ("recib", "inyec", "entreg", "combust", "gnl", "import", "cuenca", "linepack", "line-pack", "rgr",
             "transport", "balance", "perdid", "grec", "gent")


def probe_files(links):
    print("\n===== D. candidate files =====")
    cands = [h for h in links if re.search(r"\.(xlsx?|csv)(\?|$)", h, re.I)]
    print(f"{len(cands)} spreadsheet links")
    for h in cands:
        if not any(k in h.lower() or k in links[h].lower() for k in FILE_KEYS):
            continue
        r = get(h)
        if r is None or r.status_code != 200:
            print(f"-- {h}: {getattr(r, 'status_code', None)}")
            continue
        print(f"\n-- {links[h]!r} {h} bytes={len(r.content)}")
        try:
            if h.lower().endswith(".csv"):
                df = pd.read_csv(io.BytesIO(r.content), sep=None, engine="python", encoding="latin-1", nrows=40)
                print(df.head(12).to_string()[:3000])
                continue
            xl = pd.ExcelFile(io.BytesIO(r.content))
            print("   sheets:", xl.sheet_names)
            for sh in xl.sheet_names[:4]:
                df = xl.parse(sh, header=None)
                print(f"   [{sh}] shape={df.shape}")
                print(df.head(14).iloc[:, :16].to_string()[:3500])
                print("   ...last rows:")
                print(df.tail(4).iloc[:, :16].to_string()[:1500])
        except Exception as e:  # noqa: BLE001 - discovery: print and move on
            print(f"   parse error {type(e).__name__}: {e}")


def probe_se():
    print("\n===== E. Series API + CKAN =====")
    for q in ("gas inyectado", "gas entregado", "inyeccion gas", "importacion gas natural", "gnl", "gas natural licuado",
              "gas retenido", "gas venteado", "destino produccion gas", "gas aventado", "gas reinyectado"):
        r = get("https://apis.datos.gob.ar/series/api/search/", params={"q": q, "limit": 20})
        if r is None or r.status_code != 200:
            print(f"series q={q!r}: {getattr(r, 'status_code', None)}")
            continue
        hits = r.json().get("data", [])
        print(f"series q={q!r}: {len(hits)}")
        for h in hits[:12]:
            f, ds = h.get("field", {}), h.get("dataset", {})
            print(f"   {f.get('id')} | {f.get('description', '')[:70]} | {f.get('frequency')} "
                  f"{f.get('time_index_start')}..{f.get('time_index_end')} | {ds.get('title', '')[:50]}")
    for q in ("produccion gas natural", "gas inyectado", "destino de la produccion", "importacion gas", "gnl",
              "sesco", "produccion de petroleo y gas por pozo"):
        r = get("http://datos.energia.gob.ar/api/3/action/package_search", params={"q": q, "rows": 10})
        if r is None or r.status_code != 200:
            print(f"ckan q={q!r}: {getattr(r, 'status_code', None)}")
            continue
        res = r.json().get("result", {})
        print(f"\nckan q={q!r}: {res.get('count')}")
        for p in res.get("results", [])[:10]:
            print(f"  * {p.get('name')} | {p.get('title', '')[:80]}")
            for rs in p.get("resources", [])[:8]:
                print(f"      - {rs.get('name', '')[:60]} | {rs.get('format')} | {rs.get('url', '')[:150]}")


def main():
    S.get(PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)
    probe_imports()
    probe_js()
    links = probe_pages()
    probe_files(links)
    probe_se()


if __name__ == "__main__":
    main()
