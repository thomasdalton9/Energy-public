"""
Discovery for daily generation by fuel from the raw grid operators of
Panama (CND / ETESA), Costa Rica (ICE / CENCE) and Nicaragua (CNDC / ENATREL).

Round 1: fetch the operators' front pages and known report/API leads, save
every response under ca_raw/ (pushed by the workflow to the throwaway branch
ca-power-discovery-raw) and print every link / URL-like string found, so the
real pull scripts can be written against what is actually there.

Leads:
  Costa Rica: CenceWeb JSON service data/sen/json/EnergiaHorariaFuentePlanta
              (hourly energy by source and plant), CencePosdespachoNacional.jsf.
  Nicaragua:  cndc.org.ni consultas/reportesDiarios/postDespachoEnergia.php
              (post-dispatch hourly energy by plant), graficos pages.
  Panama:     cnd.com.pa informes (daily operations reports), sitr.cnd.com.pa.

Usage: python3 CENTRAL_AMERICA_POWER_DISCOVERY.py [round]
"""
import os
import re
import sys
import urllib.parse

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}
TIMEOUT = (15, 60)
RAW = "ca_raw"
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
SEEN = set()


def safe_name(url):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", url.split("://", 1)[-1])
    return s[:180]


def get(url, label=None, save=True, method="GET", **kw):
    print(f"\n{'=' * 70}\n{label or ''} {method} {url} {kw.get('params') or ''} {kw.get('data') or ''}", flush=True)
    try:
        r = SESSION.request(method, url, timeout=TIMEOUT, **kw)
    except Exception as e:  # noqa: BLE001 - discovery: report and move on
        print(f"  ERROR {type(e).__name__}: {e}", flush=True)
        return None
    ct = r.headers.get("content-type", "")
    print(f"  status={r.status_code} bytes={len(r.content)} ct={ct} final={r.url}", flush=True)
    if save:
        os.makedirs(RAW, exist_ok=True)
        name = safe_name(r.url + ("_" + urllib.parse.urlencode(kw.get("params") or {}) if kw.get("params") else ""))
        if "json" in ct and not name.endswith(".json"):
            name += ".json"
        elif "html" in ct and not name.endswith((".html", ".htm")):
            name += ".html"
        with open(os.path.join(RAW, name), "wb") as f:
            f.write(r.content)
    if "text" in ct or "json" in ct or "javascript" in ct or "xml" in ct:
        print("  head: " + r.text[:600].replace("\n", " "), flush=True)
    return r


def links(r):
    """All hrefs/srcs plus URL-ish strings (php/json/xls/api) in a page."""
    if r is None:
        return []
    text = r.text
    found = set(re.findall(r"""(?:href|src|action)\s*=\s*["']([^"'#]+)["']""", text, re.I))
    found |= set(re.findall(r"""["']([^"'\s]*(?:\.php|\.json|\.xlsx?|\.csv|\.pdf|\.jsf|/api/|/json/|\.aspx|\.ashx)[^"'\s]*)["']""", text, re.I))
    out = sorted({urllib.parse.urljoin(r.url, u.strip()) for u in found if not u.startswith(("javascript:", "mailto:"))})
    return out


def crawl(start, keywords, depth=1, limit=60):
    r = get(start, "START")
    ls = links(r)
    print(f"  {len(ls)} links:")
    for u in ls:
        print("    " + u)
    if depth <= 0:
        return
    host = urllib.parse.urlparse(start).netloc.split(".", 1)[-1]
    n = 0
    for u in ls:
        if n >= limit:
            break
        if host not in urllib.parse.urlparse(u).netloc:
            continue
        if u in SEEN or not any(k in u.lower() for k in keywords):
            continue
        if re.search(r"\.(png|jpe?g|gif|svg|ico|woff2?|ttf|css)(\?|$)", u, re.I):
            continue
        SEEN.add(u)
        n += 1
        r2 = get(u, "FOLLOW")
        if r2 is not None and "html" in r2.headers.get("content-type", ""):
            sub = links(r2)
            print(f"  {len(sub)} links:")
            for v in sub:
                print("    " + v)


def costa_rica():
    base = "https://apps.grupoice.com/CenceWeb/"
    crawl(base, ["cence", "json", "posdespacho", "genera", "data", ".js"], depth=1, limit=40)
    for path in ["CencePosdespachoNacional.jsf", "CenceMenu.jsf", "CenceDescargaDatos.jsf", "CenceIndex.jsf"]:
        r = get(base + path)
        for u in links(r):
            print("    " + u)
    for ep in ["EnergiaHorariaFuentePlanta", "GeneracionHorariaFuente", "EnergiaHorariaFuente",
               "GeneracionFuente", "EnergiaDiariaFuente", "EnergiaFuente", "DemandaMW"]:
        get(f"{base}data/sen/json/{ep}", params={"inicio": "20250105", "fin": "20250105"})
    get(f"{base}data/sen/json/EnergiaHorariaFuentePlanta", params={"inicio": "20210101", "fin": "20210102"})


def nicaragua():
    base = "https://www.cndc.org.ni/"
    crawl(base, ["consulta", "reporte", "despacho", "genera", "informe", "diario", "grafic", "operac", "estadist"],
          depth=1, limit=60)
    for u, params in [
        (base + "consultas/reportesDiarios/postDespachoEnergia.php", {"fecha": "05/01/2025"}),
        (base + "consultas/reportesDiarios/postDespachoEnergia.php", {"d": "05/01/2025"}),
        (base + "consultas/reportesDiarios/postDespachoEnergia.php", None),
        (base + "graficos/graficaGeneracion_Tipo_TReal.php", None),
        (base + "graficos/generacion_tipo_tiemporeal.php", None),
    ]:
        get(u, params=params)


def panama():
    crawl("https://www.cnd.com.pa/", ["informe", "reporte", "despacho", "genera", "operac", "diario", "estadist",
                                       "documento", "categoria"], depth=1, limit=60)
    for u in ["https://sitr.cnd.com.pa/", "https://sitr.cnd.com.pa/m/pub/gen.html",
              "https://www.cnd.com.pa/index.php/informes", "https://www.cnd.com.pa/informes"]:
        r = get(u)
        for v in links(r):
            print("    " + v)


if __name__ == "__main__":
    which = sys.argv[1:] or ["cr", "ni", "pa"]
    for w in which:
        print(f"\n\n######## {w} ########", flush=True)
        try:
            {"cr": costa_rica, "ni": nicaragua, "pa": panama}[w]()
        except Exception as e:  # noqa: BLE001
            print(f"FAILED {w}: {type(e).__name__}: {e}", flush=True)
