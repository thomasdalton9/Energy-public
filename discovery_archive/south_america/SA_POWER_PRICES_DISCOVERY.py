"""
Discovery probe (round 1) for the South America daily wholesale power price pull
(south_america/SA_POWER_PRICES_DAILY.py). One-off; run by the manual workflow
discovery_archive/workflows/sa_power_prices_discovery.yml.

Probes, printing a short sample of each response:
  Brazil   CCEE open-data portal (CKAN) - PLD datasets; BCB PTAX (SGS series 1, Olinda PTAX)
  Colombia XM servapibi metric list (price metrics), PrecBolsNaci hourly for one week; TRM (datos.gov.co, Banrep)
  Peru     COES costo marginal pages; BCRP exchange rate series
  Argentina CAMMESA post-operative mdb tables with prices; other nemos; BCRA exchange-rate API
  Chile    cne.cl media search for marginal-cost / market-price files
  Uruguay  ADME site links for the spot price
  Bolivia  CNDC categories / rt endpoints for marginal costs
  Ecuador  CENACE (TLS-verified only; reports if the chain fails)
"""
import csv
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import traceback
import zipfile

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
ONLY = set(sys.argv[1:])


def head(title):
    print("\n" + "=" * 100 + f"\n{title}\n" + "=" * 100, flush=True)


def show(r, n=1500):
    print(f"  -> {r.status_code} {r.headers.get('content-type')} len={len(r.content)} url={r.url[:250]}")
    print("  " + r.text[:n].replace("\n", "\n  "))


def get(url, n=1500, **kw):
    kw.setdefault("timeout", 60)
    kw.setdefault("headers", UA)
    try:
        r = requests.get(url, **kw)
        show(r, n)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  !! {url[:200]}: {type(e).__name__}: {str(e)[:300]}")
        return None


def post(url, n=1500, **kw):
    kw.setdefault("timeout", 60)
    kw.setdefault("headers", UA)
    try:
        r = requests.post(url, **kw)
        show(r, n)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  !! {url[:200]}: {type(e).__name__}: {str(e)[:300]}")
        return None


def links(html, pat):
    found = sorted(set(re.findall(r"""(?:href|src|url)\s*[=:]\s*["']([^"']+)["']""", html, re.I)))
    return [u for u in found if re.search(pat, u, re.I)]


def run(name, fn):
    if ONLY and name not in ONLY:
        return
    head(name)
    try:
        fn()
    except Exception:  # noqa: BLE001
        traceback.print_exc()


# ------------------------------------------------------------------ Brazil
def brazil():
    base = "https://dadosabertos.ccee.org.br/api/3/action/"
    r = get(base + "package_search?q=pld&rows=50", n=300)
    if r is not None and r.ok:
        for p in r.json()["result"]["results"]:
            print(f"  PKG {p['name']}: {p.get('title')}")
            for res in p.get("resources", [])[:40]:
                print(f"     res {res.get('id')} {res.get('name')} {res.get('format')} "
                      f"ds={res.get('datastore_active')} {str(res.get('url'))[:140]}")
    for pkg in ("pld_horario", "pld_media_diaria"):
        r = get(base + f"package_show?id={pkg}", n=200)
    # try datastore on the newest pld_horario resource
    r = requests.get(base + "package_show?id=pld_horario", headers=UA, timeout=60)
    if r.ok:
        res = r.json()["result"]["resources"]
        for x in res[-2:]:
            print(f"  datastore_search {x['name']}")
            get(base + f"datastore_search?resource_id={x['id']}&limit=5", n=1500)
            get(base + f"datastore_search?resource_id={x['id']}&limit=3&sort=_id desc", n=1500)
        first = res[0]
        print(f"  first resource {first['name']} url={first.get('url')}")
        try:
            rr = requests.get(first["url"], headers=UA, timeout=120)
            print(f"  csv {rr.status_code} {len(rr.content)} bytes")
            print("  " + rr.content[:600].decode("latin-1").replace("\n", "\n  "))
        except Exception as e:  # noqa: BLE001
            print("  !!", e)
    # PTAX
    get("https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados?formato=json&dataInicial=01/09/2026&dataFinal=30/09/2026",
        n=600)
    get("https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarPeriodo(dataInicial=@dataInicial,"
        "dataFinalCotacao=@dataFinalCotacao)?@dataInicial='09-01-2026'&@dataFinalCotacao='09-10-2026'&$format=json",
        n=600)


# ------------------------------------------------------------------ Colombia
def colombia():
    xm = "https://servapibi.xm.com.co/"
    r = post(xm + "lists", n=200, json={"MetricId": "ListadoMetricas"}, headers={"Connection": "close"})
    if r is not None and r.ok:
        for it in r.json().get("Items", []):
            for e in it.get("ListEntities", []):
                v = e.get("Values", {})
                blob = json.dumps(v, ensure_ascii=False)
                if re.search(r"prec|escas|cost|marg", blob, re.I):
                    print("  METRIC", blob[:400])
    body = {"MetricId": "PrecBolsNaci", "Entity": "Sistema", "StartDate": "2026-09-01", "EndDate": "2026-09-07"}
    post(xm + "hourly", n=1500, json=body, headers={"Connection": "close"})
    body = {"MetricId": "PrecBolsNaci", "Entity": "Sistema", "StartDate": "2021-01-01", "EndDate": "2021-01-31"}
    r = post(xm + "hourly", n=300, json=body, headers={"Connection": "close"})
    if r is not None and r.ok:
        print("  2021-01 items:", len(r.json().get("Items", [])))
    for m in ("PPPrecBolsNaci", "PrecEsca", "PrecEscaAct", "PrecEscaMarg", "PrecEscaPon", "MaxPrecOferNal"):
        print(f"  daily {m}")
        post(xm + "daily", n=700, json={"MetricId": m, "Entity": "Sistema", "StartDate": "2026-09-01",
                                         "EndDate": "2026-09-05"}, headers={"Connection": "close"})
    for m in ("PrecEsca", "PrecEscaAct"):
        print(f"  monthly {m}")
        post(xm + "monthly", n=700, json={"MetricId": m, "Entity": "Sistema", "StartDate": "2026-01-01",
                                           "EndDate": "2026-09-30"}, headers={"Connection": "close"})
    # TRM
    get("https://www.datos.gov.co/resource/32sa-8pi3.json?$where=vigenciadesde>='2026-09-01T00:00:00'"
        "&$order=vigenciadesde&$limit=5", n=800)
    get("https://www.datos.gov.co/resource/mcec-87by.json?$limit=3", n=600)
    get("https://totoro.banrep.gov.co/nsi-jax-ws/rest/data/ESTAT,DF_TRM_DAILY_HIST,1.0/all/ALL/"
        "?startPeriod=2026-09-01&endPeriod=2026-09-10&dimensionAtObservation=TIME_PERIOD&detail=full", n=800)
    get("https://www.banrep.gov.co/es/estadisticas/trm", n=300)


# ------------------------------------------------------------------ Peru
def peru():
    for u in ("https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales",
              "https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/index",
              "https://www.coes.org.pe/Portal/portalinformacion/costosmarginales",
              "https://www.coes.org.pe/Portal/portalinformacion/costomarginal",
              "https://www.coes.org.pe/Portal/portalinformacion"):
        r = get(u, n=200)
        if r is not None and r.ok:
            for line in sorted(set(re.findall(r"""['"](/?Portal/[A-Za-z/]+|[a-z]+/[a-z]+)['"]""", r.text)))[:60]:
                print("    ref", line)
            for m in re.findall(r"(?:url|action)\s*[:=]\s*['\"]([^'\"]+)['\"]", r.text)[:40]:
                print("    ajax", m)
            for m in re.findall(r"<select[^>]*id=['\"]([^'\"]+)['\"]", r.text)[:20]:
                print("    select", m)
            i = r.text.lower().find("marginal")
            if i >= 0:
                print("    ctx:", r.text[max(0, i - 300): i + 500].replace("\n", " "))
    sess = requests.Session()
    sess.headers.update({**UA, "X-Requested-With": "XMLHttpRequest"})
    for u, data in (("https://www.coes.org.pe/Portal/portalinformacion/costosmarginales",
                     {"fechaInicial": "01/09/2026", "fechaFinal": "01/09/2026"}),
                    ("https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/consulta",
                     {"fechaInicial": "01/09/2026", "fechaFinal": "01/09/2026", "barras": ""}),
                    ("https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/grafico",
                     {"fechaInicial": "01/09/2026", "fechaFinal": "01/09/2026", "barras": ""})):
        print("  POST", u, data)
        try:
            r = sess.post(u, data=data, timeout=60)
            show(r, 1500)
        except Exception as e:  # noqa: BLE001
            print("  !!", e)
    # generation page JSON keys (maybe it carries a CMg series too)
    try:
        r = sess.post("https://www.coes.org.pe/Portal/portalinformacion/generacion",
                      data={"fechaInicial": "01/09/2026", "fechaFinal": "01/09/2026", "indicador": 0}, timeout=60)
        print("  generacion keys:", list(r.json().keys()))
    except Exception as e:  # noqa: BLE001
        print("  !!", e)
    get("https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PD04640PD/json/2026-09-01/2026-09-10", n=800)
    get("https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PD04638PD/json/2026-09-01/2026-09-10", n=800)


# ------------------------------------------------------------------ Argentina
def argentina():
    look = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
    att = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
    fmt = "%Y-%m-%dT%H:%M:%S.000Z"
    s = requests.Session()
    s.headers.update({"User-Agent": "gas-demand-scripts/1.0"})
    r = s.get(look, params={"fechadesde": "2026-09-01T00:00:00.000Z", "fechahasta": "2026-09-02T00:00:00.000Z",
                            "nemo": "PARTE_POST_OPERATIVO"}, timeout=60)
    print("  docs", r.status_code, r.text[:600])
    docs = r.json()
    doc = docs[0]
    a = next(x for x in doc["adjuntos"] if x["id"].endswith(".zip"))
    z = s.get(att, params={"docId": doc["id"], "attachmentId": a["id"], "nemo": "PARTE_POST_OPERATIVO"}, timeout=120)
    print("  zip", z.status_code, len(z.content))
    zf = zipfile.ZipFile(io.BytesIO(z.content))
    print("  members", zf.namelist())
    name = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(zf.read(name))
        path = f.name
    tables = subprocess.run(["mdb-tables", "-1", path], capture_output=True, text=True).stdout.split()
    print("  tables:", tables)
    for t in tables:
        out = subprocess.run(["mdb-export", path, t], capture_output=True, text=True).stdout
        hdr = out.splitlines()[0] if out else ""
        if re.search(r"PREC|COST|MARG|CMG|SPOT|MONOM", t + " " + hdr, re.I):
            rows = out.splitlines()
            print(f"  TABLE {t} rows={len(rows) - 1}\n    " + "\n    ".join(rows[:8]))
    # other nemos that may carry prices
    for nemo in ("PRECIOS_MERCADO", "PRECIO_SPOT", "PRECIOS", "COSTOS_MARGINALES", "INFORME_MENSUAL",
                 "INFORME_MENSUAL_PRINCIPALES_VARIABLES", "PRECIOS_MONOMICOS", "INF_SINTESIS_MENSUAL",
                 "TRANSACCIONES_ECONOMICAS", "DTE", "PRECIOS_ESTACIONALES"):
        try:
            r = s.get(look, params={"fechadesde": "2026-06-01T00:00:00.000Z", "fechahasta": "2026-09-30T00:00:00.000Z",
                                    "nemo": nemo}, timeout=60)
            j = r.json() if r.ok else None
            n = len(j) if isinstance(j, list) else j
            print(f"  nemo {nemo}: {r.status_code} n={str(n)[:100]}")
            if isinstance(j, list) and j:
                for d in j[:3]:
                    print("     ", json.dumps({k: d.get(k) for k in ("id", "titulo", "fecha", "nemo")},
                                             ensure_ascii=False)[:200],
                          [x.get("id") for x in d.get("adjuntos", [])][:6])
        except Exception as e:  # noqa: BLE001
            print(f"  nemo {nemo}: !! {e}")
    # CAMMESA web data APIs
    for u in ("https://api.cammesa.com/demanda-svc/demanda/ObtieneDemandaYTemperaturaRegion?id_region=1002",
              "https://cammesaweb.cammesa.com/precios-y-costos/",
              "https://cammesaweb.cammesa.com/"):
        r = get(u, n=300)
        if r is not None and r.ok and "html" in str(r.headers.get("content-type")):
            for l_ in links(r.text, r"prec|cost|monom|spot|marg")[:40]:
                print("    link", l_)
    # BCRA
    get("https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones/USD?fechadesde=2026-09-01"
        "&fechahasta=2026-09-10", n=900)
    get("https://api.bcra.gob.ar/estadisticas/v3.0/monetarias/5?desde=2026-09-01&hasta=2026-09-10", n=900)
    get("https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/5?desde=2026-09-01&hasta=2026-09-10", n=900)


# ------------------------------------------------------------------ Chile
def chile():
    media = "https://www.cne.cl/wp-json/wp/v2/media"
    seen = set()
    for q in ("marginal", "costo marginal", "Costos_Marginales", "CMg", "precio medio", "PMM", "precio de nudo",
              "precio", "Reporte Mensual", "reporte_mensual", "spot"):
        try:
            r = requests.get(media, params={"search": q, "per_page": 50, "_fields": "date,modified,source_url,title"},
                             headers=UA, timeout=60)
            items = r.json() if r.ok else []
            print(f"  search {q!r}: {r.status_code} {len(items)}")
            for i in items:
                u = i.get("source_url", "")
                if u not in seen:
                    seen.add(u)
                    print(f"     {i.get('date', '')[:10]} {u}")
        except Exception as e:  # noqa: BLE001
            print("  !!", e)
    for u in ("https://www.cne.cl/estadisticas/electricidad/", "https://www.cne.cl/normativas/electrica/precios/",
              "https://www.cne.cl/tarificacion/electricidad/precio-medio-de-mercado/",
              "https://www.cne.cl/nuestros-servicios/reportes/informacion-y-estadisticas/"):
        r = get(u, n=100)
        if r is not None and r.ok:
            for l_ in links(r.text, r"margin|cmg|precio|pmm|nudo|reporte")[:60]:
                print("    link", l_)


# ------------------------------------------------------------------ Uruguay
def uruguay():
    for u in ("https://pronos.adme.com.uy/", "https://adme.com.uy/", "https://www.adme.com.uy/",
              "https://adme.com.uy/mmee/", "https://pronos.adme.com.uy/gpf.php"):
        r = get(u, n=200)
        if r is not None and r.ok:
            for l_ in links(r.text, r"spot|precio|cmg|marginal|costo|php|cgi")[:60]:
                print("    link", l_)
    for u in ("https://pronos.adme.com.uy/cgi-bin/series_spot.cgi", "https://pronos.adme.com.uy/spot.php",
              "https://pronos.adme.com.uy/pspot.php", "https://pronos.adme.com.uy/cmg.php",
              "https://adme.com.uy/mmee/infspot.php", "https://adme.com.uy/mmee/preciospot.php"):
        get(u, n=500)


# ------------------------------------------------------------------ Bolivia / Ecuador
def bolivia():
    api = "https://www.cndc.bo/wp-json/cndc/v1/"
    for tipo in ("estadisticas", "operacion", "mercado", "informes", ""):
        r = get(api + "estadisticas/categorias" + (f"?tipo={tipo}" if tipo else ""), n=200)
        if r is not None and r.ok:
            txt = r.text
            for m in re.finditer(r"\{[^{}]*(?:[Mm]arginal|[Cc]osto|[Pp]recio|[Tt]arifa)[^{}]*\}", txt):
                print("    cat", m.group(0)[:300])
    for ep in ("rt/costomarginal", "rt/costo_marginal", "rt/costos-marginales", "rt/cmg", "rt/costos",
               "rt/precio", "rt/demanda"):
        get(api + ep + "?fecha=2026-09-01", n=500)
    r = get(api, n=200)
    if r is not None and r.ok:
        try:
            routes = list(r.json().get("routes", {}).keys())
            print("  routes:", routes[:200])
        except Exception:  # noqa: BLE001
            pass


def ecuador():
    get("https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm", n=200)
    get("https://www.cenace.gob.ec/", n=200)


run("brazil", brazil)
run("colombia", colombia)
run("peru", peru)
run("argentina", argentina)
run("chile", chile)
run("uruguay", uruguay)
run("bolivia", bolivia)
run("ecuador", ecuador)
print("\nDONE")
