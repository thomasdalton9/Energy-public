"""One-off probe 3: COES new-project pages (Peru) - find the AJAX endpoints behind the tables."""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
S = requests.Session()
S.headers.update(H)


def get(u, **kw):
    try:
        r = S.get(u, timeout=(10, 120), **kw)
        print(f"  GET {u[:140]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content):,}B", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  GET {u[:140]} -> {type(e).__name__}: {str(e)[:120]}", flush=True)
        return None


def links(r, pat, n=40):
    if r is None or not r.ok:
        return []
    out = []
    for href, text in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        if re.search(pat, href + " " + t, re.I):
            out.append((href, t[:90]))
    for h, t in out[:n]:
        print(f"    link: {t} | {h[:200]}")
    return out


def peek(r, label=""):
    if r is None or not r.ok:
        return
    c = r.content
    try:
        if c[:2] == b"PK" or c[:4] == b"\xd0\xcf\x11\xe0":
            xl = pd.ExcelFile(io.BytesIO(c))
            print(f"    {label} sheets: {xl.sheet_names[:20]}")
            for sh in xl.sheet_names[:6]:
                d = pd.read_excel(xl, sh, header=None, nrows=25)
                print(f"    --- {sh} {d.shape}\n{d.dropna(how='all').head(14).to_string(max_colwidth=28)[:2500]}")
        elif b"," in c[:2000] or b";" in c[:2000]:
            d = pd.read_csv(io.BytesIO(c), sep=None, engine="python", nrows=10, encoding="latin-1")
            print(f"    {label} csv columns {list(d.columns)}\n{d.head(5).to_string(max_colwidth=25)[:2000]}")
    except Exception as e:  # noqa: BLE001
        print(f"    peek failed {type(e).__name__}: {e}")



BASE = "https://www.coes.org.pe"
for u in ("/Portal/Planificacion/NuevosProyectos/EstudiosPO", "/Portal/Planificacion/NuevosProyectos/OperacionComercialUnidades"):
    r = get(BASE + u)
    if r is None or not r.ok:
        continue
    for sc in re.findall(r"<script[^>]*>(.*?)</script>", r.text, re.S):
        if re.search(r"ajax|\.post|\.get|url", sc, re.I) and len(sc) > 50:
            print("    inline script:", re.sub(r"\s+", " ", sc)[:2500])
    for sc in re.findall(r'<script[^>]+src="([^"]+)"', r.text):
        if not re.search(r"jquery|bootstrap|modernizr|google|analytics", sc, re.I):
            js = get(BASE + sc if sc.startswith("/") else sc)
            if js is not None and js.ok:
                body = re.sub(r"[ \\t]+", " ", js.text)[:3000]
                print(f"      js {sc}: {body}")
    for m in re.findall(r'<(?:select|input)[^>]+(?:id|name)="([^"]+)"', r.text)[:30]:
        print("    field:", m)
for path in ("/Portal/Planificacion/NuevosProyectos/EstudiosPO/Lista", "/Portal/Planificacion/NuevosProyectos/ListaEstudiosPO",
             "/Portal/Planificacion/NuevosProyectos/EstudiosPOLista"):
    for meth in ("get", "post"):
        try:
            r = getattr(S, meth)(BASE + path, timeout=(10, 60))
            print(f"  {meth.upper()} {path} -> {r.status_code} {r.headers.get('content-type', '')[:30]} {len(r.content)}B {r.text[:300]!r}")
        except Exception as e:  # noqa: BLE001
            print(f"  {meth} {path} -> {type(e).__name__}")
