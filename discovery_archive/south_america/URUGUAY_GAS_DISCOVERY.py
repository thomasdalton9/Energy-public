"""
One-off probe: what Uruguay natural gas demand data (by sector / tariff) is
downloadable, and in what layout.

Round 1 (pages + CKAN): catalogodatos.gub.uy has nothing for 'gas natural';
ben.miem.gub.uy no longer resolves; the URSEA statistics page has no gas
files. The MIEM 'Series estadisticas de gas natural' page offers one zip
of all attachments (download/node/field_documento/3815) plus a monthly
visualiser (visualpeb.miem.gub.uy/visualPEB/gas_natural).

Round 2 (this version): download and dump the MIEM zips, look at the
visualPEB page / scripts, and list BEN (balance) links on the MIEM data
pages. Runs in GitHub Actions only.
"""
import io
import re
import zipfile
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 60)
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 32)

ZIPS = [
    "https://www.gub.uy/ministerio-industria-energia-mineria/download/node/field_documento/3815",  # gas series
    "https://www.gub.uy/ministerio-industria-energia-mineria/download/node/field_documento/3876",  # discontinued
]
PAGES = [
    "https://visualpeb.miem.gub.uy/visualPEB/gas_natural",
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos",
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos-abiertos",
    "https://www.gub.uy/ministerio-industria-energia-mineria/tematica/planificacion-estadistica-balance",
    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/estadisticas",
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
        out(raw.head(25).iloc[:, :15].to_string())
        out("  ... last rows:")
        out(raw.tail(6).iloc[:, :15].to_string())


def handle_file(url, label=""):
    if url in seen or len(seen) > 40:
        return
    seen.add(url)
    out(f"\n##### FILE {label!r} {url}")
    r = get(url)
    if r is None or r.status_code != 200:
        return
    c = r.content
    if c[:2] == b"PK":
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


def main():
    for z in ZIPS:
        handle_file(z, "zip")
    for page in PAGES:
        out(f"\n########## PAGE {page}")
        r = get(page)
        if r is None or r.status_code != 200:
            continue
        if "visualpeb" in page:
            out(r.text[:3000])
            for src in re.findall(r'<script[^>]+src="([^"]+)"', r.text, re.I):
                js = get(urljoin(page, src))
                if js is not None:
                    found = sorted(set(re.findall(r"[\"'](/?(?:api|data|visualPEB|static)[^\"']{2,120})[\"']", js.text)))
                    out(f"  script {src}: urls {found[:60]}")
        ls = links(r.text, page)
        for u, t in ls:
            if FILE_RE.search(u) or re.search(r"gas|balance|serie|consumo|visualiz|sector|energ", u + " " + t, re.I):
                out(f"  LINK {t[:90]!r} -> {u}")
        iframes = re.findall(r'<iframe[^>]+src="([^"]+)"', r.text, re.I)
        out(f"  iframes: {iframes}")


if __name__ == "__main__":
    main()
