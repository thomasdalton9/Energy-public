"""One-off probe 2: UPME active-registrations workbook layout; COES pre-operability / commercial-operation pages (Peru)."""
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



print("=================== COLOMBIA: UPME active registrations workbook")
r = get("https://docs.upme.gov.co/SIMEC/Energia%20Electrica/Informes_Registro_Proyectos_Generacion/"
        "Informe_registros_activos_de_proyectos_generacion_electrica_agosto_2026.xlsx")
if r is not None and r.ok:
    xl = pd.ExcelFile(io.BytesIO(r.content))
    print("  sheets", xl.sheet_names)
    for sh in xl.sheet_names[:4]:
        d = pd.read_excel(xl, sh, header=None)
        print(f"  --- {sh} {d.shape}")
        print(d.dropna(how="all").head(12).to_string(max_colwidth=30)[:3500])
        for c in d.columns[:25]:
            v = d[c].dropna().astype(str)
            if 2 < v.nunique() < 25:
                print(f"    col {c} values: {v.value_counts().head(12).to_dict()}")

print("=================== PERU: COES new-project pages")
for u in ("https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos/EstudiosPO",
          "https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos/OperacionComercialUnidades",
          "https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos/ConclusionOperacion"):
    r = get(u)
    if r is None or not r.ok:
        continue
    links(r, r"xlsx|xls|export|descarg|excel|listar|lista", 20)
    for m in re.findall(r"(?:url\s*:\s*|\$\.(?:post|get|ajax)\(\s*)['\"]([^'\"]+)['\"]", r.text)[:20]:
        print("    ajax:", m)
    for m in re.findall(r"<form[^>]*action=\"([^\"]+)\"", r.text)[:5]:
        print("    form:", m)
    tabs = re.findall(r"<table.*?</table>", r.text, re.S)
    first = re.sub(r"<[^>]+>|\\s+", " ", tabs[0])[:800] if tabs else ""
    print(f"    {len(tabs)} tables; first table text: {first}")
    for sc in re.findall(r'<script[^>]+src="([^"]+)"', r.text):
        if "NuevosProyectos" in sc or "Planificacion" in sc or "estudio" in sc.lower():
            js = get("https://www.coes.org.pe" + sc if sc.startswith("/") else sc)
            if js is not None and js.ok:
                for m in re.findall(r"['\"](/?Portal/[^'\"]+|[A-Za-z]+/[A-Za-z]+(?:Listar|Lista|Export|Descargar)[^'\"]*)['\"]", js.text)[:30]:
                    print("      js url:", m)
                print("      js head:", js.text[:1500].replace("\n", " "))
