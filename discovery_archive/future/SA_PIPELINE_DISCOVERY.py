"""One-off probe: machine-readable generation project pipelines for Chile, Colombia, Peru (for future/SA_FUTURE.py)."""
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


print("=================== CHILE: CNE media library (wp-json) - projects in construction / declared in construction")
for q in ("construccion", "construcción", "proyectos", "Reporte Mensual ERNC", "en construccion", "capacidad instalada"):
    r = get("https://www.cne.cl/wp-json/wp/v2/media", params={"search": q, "per_page": 30, "orderby": "date"})
    if r is not None and r.ok:
        for m in r.json()[:30]:
            print(f"    {m.get('date', '')[:10]} {m.get('title', {}).get('rendered', '')[:70]} | {m.get('source_url')}")
cands = []
r = get("https://www.cne.cl/wp-json/wp/v2/media", params={"search": "construcci", "per_page": 50, "orderby": "date"})
if r is not None and r.ok:
    cands = [m["source_url"] for m in r.json() if re.search(r"\.xlsx?$", m.get("source_url", ""), re.I)]
for u in cands[:3]:
    peek(get(u), u.rsplit("/", 1)[-1])
for u in ("https://www.cne.cl/normativas/electrica/", "https://www.cne.cl/estadisticas/electricidad/",
          "https://www.cne.cl/tarificacion/electricidad/proyectos-en-construccion/",
          "http://energiaabierta.cl/?s=construccion", "https://energiaabierta.cl/categorias-estadistica/electricidad/",
          "https://www.coordinador.cl/desarrollo/documentos/proyectos-en-construccion/",
          "https://www.coordinador.cl/desarrollo/graficos/proyectos-en-construccion/"):
    links(get(u), r"construcci|proyecto|\.xlsx|declarad", 25)

print("=================== COLOMBIA: UPME project register / XM expansion")
for q in ("registro proyectos generacion", "proyectos de generacion inscritos", "UPME proyectos generacion",
          "proyectos generacion energia electrica", "subasta", "expansion generacion"):
    r = get("https://www.datos.gov.co/api/catalog/v1", params={"q": q, "limit": 10})
    if r is not None and r.ok:
        for x in r.json().get("results", [])[:10]:
            res = x.get("resource", {})
            print(f"    {res.get('id')} | {res.get('name', '')[:90]} | {res.get('updatedAt', '')[:10]} | {res.get('type')}")
for u in ("https://www1.upme.gov.co/siel/Pages/Registro-de-proyectos-de-generacion.aspx",
          "https://www1.upme.gov.co/Paginas/Registro-de-proyectos-de-generacion.aspx",
          "https://www.upme.gov.co/registro-de-proyectos-de-generacion/",
          "https://www1.upme.gov.co/Energia_electrica/Paginas/Registro-de-Proyectos-de-Generacion.aspx",
          "https://www.upme.gov.co/energia-electrica/registro-de-proyectos-de-generacion/",
          "https://www.xm.com.co/transmision/proyectos-de-expansion",
          "https://www.xm.com.co/generaci%C3%B3n/proyectos-de-generaci%C3%B3n"):
    found = links(get(u), r"registro|proyecto|\.xlsx|\.xls|inscri", 30)
    for h, _ in found:
        if re.search(r"\.xlsx?($|\?)", h, re.I):
            url = h if h.startswith("http") else re.match(r"https?://[^/]+", u).group(0) + ("" if h.startswith("/") else "/") + h
            peek(get(url), h.rsplit("/", 1)[-1])
            break

print("=================== PERU: COES new projects / MINEM / OSINERGMIN")
for u in ("https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos",
          "https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos/OperacionComercial",
          "https://www.coes.org.pe/Portal/Planificacion/NuevosProyectos/EstudiosPreOperatividad",
          "https://www.coes.org.pe/Portal/Planificacion/Estudios/EstudiosPreOperatividad",
          "https://www.coes.org.pe/Portal/Planificacion/",
          "https://www.gob.pe/institucion/minem/informes-publicaciones",
          "https://www.minem.gob.pe/_estadistica.php?idSector=6",
          "https://www.osinergmin.gob.pe/seccion/institucional/regulacion-tarifaria/publicaciones/proyectos-generacion",
          "https://www.osinergmin.gob.pe/empresas/electricidad/proyectos/generacion"):
    links(get(u), r"proyect|operaci|pre.?operat|\.xlsx|\.xls|cartera|generaci", 30)
r = get("https://www.datosabiertos.gob.pe/api/3/action/package_search", params={"q": "proyectos generacion electrica", "rows": 10})
if r is not None and r.ok and "json" in r.headers.get("content-type", ""):
    for p in r.json()["result"]["results"][:10]:
        print("   pkg:", p["title"][:80], [(x.get("format"), x.get("url")) for x in p.get("resources", [])][:3])
