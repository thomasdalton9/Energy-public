"""
One-off probe: what Uruguay natural gas demand data (by sector / tariff) is
downloadable, and in what layout.

Candidates:
  - MIEM / DNE 'Series estadisticas de gas natural' (billing by tariff per
    distributor, customers by tariff, prices, imports) on gub.uy
  - MIEM discontinued series page
  - catalogodatos.gub.uy (CKAN open data) - search 'gas natural'
  - ben.miem.gub.uy (Balance Energetico Nacional) - annual by sector
  - URSEA statistics pages
Lists links on each page, downloads spreadsheet/zip files and dumps every
sheet's shape, first rows and last rows. Runs in GitHub Actions only.
"""
import io
import json
import re
import sys
import zipfile
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 60)
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 40)

PAGES = [
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/estadisticas/series-estadisticas-gas-natural",
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/series-estadisticas-gas-natural",
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/estadisticas/series-estadisticas-discontinuadas",
    "https://ben.miem.gub.uy/",
    "https://ben.miem.gub.uy/descargas.php",
    "https://ben.miem.gub.uy/series.php",
    "https://www.gub.uy/unidad-reguladora-servicios-energia-agua/datos-y-estadisticas",
    "https://www.gub.uy/unidad-reguladora-servicios-energia-agua/datos-y-estadisticas/estadisticas",
]
FILE_RE = re.compile(r"\.(zip|xlsx|xls|csv|ods)(\?|$)", re.I)
seen = set()


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
        return r
    except requests.RequestException as e:
        out(f"GET {url} -> ERROR {e}")
        return None


def dump_sheets(content, name):
    try:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:
        out(f"  cannot read {name} as excel: {e}")
        try:
            df = pd.read_csv(io.BytesIO(content), sep=None, engine="python", encoding="latin-1")
            out(df.head(15).to_string())
            out(df.tail(5).to_string())
        except Exception as e2:
            out(f"  not csv either: {e2}")
        return
    for sn, raw in sheets.items():
        raw = raw.dropna(how="all").dropna(axis=1, how="all")
        out(f"  === {name} :: sheet {sn!r} shape {raw.shape}")
        out(raw.head(22).iloc[:, :16].to_string())
        out("  ... last rows:")
        out(raw.tail(6).iloc[:, :16].to_string())


def handle_file(url, label=""):
    if url in seen or len(seen) > 40:
        return
    seen.add(url)
    out(f"\n##### FILE {label!r} {url}")
    r = get(url)
    if r is None or r.status_code != 200:
        return
    c = r.content
    if c[:2] == b"PK" and (url.lower().split("?")[0].endswith(".zip") or b"xl/" not in c[:2000]):
        try:
            z = zipfile.ZipFile(io.BytesIO(c))
            names = z.namelist()
            if any(n.startswith("xl/") for n in names):
                dump_sheets(c, url)
                return
            out(f"  zip members: {names}")
            for n in names:
                if FILE_RE.search(n):
                    dump_sheets(z.read(n), n)
            return
        except zipfile.BadZipFile:
            pass
    dump_sheets(c, url)


def links(html, base):
    res = []
    for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", txt)).strip()
        res.append((urljoin(base, href.replace("&amp;", "&")), t))
    return res


def ckan():
    for base in ["https://catalogodatos.gub.uy/api/3/action/package_search"]:
        for q in ["gas natural", "gas natural miem", "balance energetico"]:
            r = get(base, params={"q": q, "rows": 30})
            if r is None or r.status_code != 200:
                continue
            try:
                res = r.json()["result"]["results"]
            except Exception as e:
                out(f"  bad json: {e} {r.text[:300]}")
                continue
            for p in res:
                out(f"  PKG {p.get('name')} | {p.get('title')} | org={p.get('organization', {}).get('name')}")
                for rs in p.get("resources", []):
                    out(f"      RES {rs.get('format')} | {rs.get('name')} | {rs.get('url')} | mod={rs.get('last_modified') or rs.get('created')}")


def main():
    out("########## CKAN catalogodatos.gub.uy")
    ckan()
    for page in PAGES:
        out(f"\n########## PAGE {page}")
        r = get(page)
        if r is None or r.status_code != 200:
            continue
        ls = links(r.text, page)
        interesting = [(u, t) for u, t in ls if FILE_RE.search(u) or re.search(
            r"gas|tarifa|factur|import|balance|serie|descarg|estad|consumo|power ?bi|tablero|visualiz", u + " " + t, re.I)]
        for u, t in interesting:
            out(f"  LINK {t[:90]!r} -> {u}")
        iframes = re.findall(r'<iframe[^>]+src="([^"]+)"', r.text, re.I)
        out(f"  iframes: {iframes}")
        for u, t in ls:
            if FILE_RE.search(u) and re.search(r"gas|GN|factur|tarif|import|client|precio|ben|serie", u + t, re.I):
                handle_file(u, t)


if __name__ == "__main__":
    main()
