"""
Round 2 (Oct-2026), Dominican Republic:
  oc2   OC-SENI chart web service (apps.oc.org.do/wsOCWebsiteChart/Service.asmx),
        seen in the www.oc.org.do/Servicios/Reporte page script: print the page
        script around each call, then try the calls for a few dates; and the
        DNN / Bring2mind document folders of Informes-Historicos (daily operation
        reports). Saves pages and answers under probe_files/.
  mem   MEM (mem.gob.do) generation bulletins / performance reports: save the
        category pages, list their files.
Usage: python3 CARIBBEAN_POWER_GAS_PROBE2.py oc2|mem
"""

print("STARTING", flush=True)

import os
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, save=None, **kw):
    try:
        r = S.get(url, timeout=60, **kw)
        print(f"[{r.status_code}] {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        if save:
            with open(os.path.join(OUT, save), "wb") as f:
                f.write(r.content)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


WS = "https://apps.oc.org.do/wsOCWebsiteChart/Service.asmx/"


def oc2():
    r = get("https://www.oc.org.do/Servicios/Reporte", save="oc_reporte.html")
    if r is not None:
        t = r.text
        for m in re.finditer(r"wsOCWebsiteChart", t):
            print("---- context\n", t[max(0, m.start() - 600): m.start() + 600], flush=True)
    # page's scripts
    for js in re.findall(r"<script[^>]+src=[\"']([^\"']+)", r.text if r is not None else "")[:40]:
        if "oc" in js.lower() and not re.search(r"jquery|bootstrap|dnn\.|telerik", js, re.I):
            print("script", urljoin("https://www.oc.org.do/", js))
    calls = [
        "GetGeneracionTotalJSon?Fecha=2026-09-28",
        "GetGeneracionTotalJSon?Fecha=28/09/2026",
        "GetGeneracionPorcentajeJSon?Fecha=2026-09-28&TIPO=ALL",
        "GetGeneracionPorcentajeJSon?Fecha=28/09/2026&TIPO=ALL",
        "GetGeneracionRenovableNotRenovableJSon",
        "GetGeneracionAcumuladaJSon?Filtro=Dx&Desde=2026-09-01&Hasta=2026-09-28",
        "GetGeneracionAcumuladaJSon?Filtro=Mx&Desde=2021-01-01&Hasta=2026-09-28",
        "GetGeneracionAcumuladaJSon?Filtro=Ax&Desde=2021-01-01&Hasta=2026-09-28",
        "GetGeneracionAcumuladaJSon?Filtro=Dx&Desde=01/09/2026&Hasta=28/09/2026",
        "GetGeneracionAcumuladaJSon?Filtro=Dx&Desde=2021-01-01&Hasta=2021-01-10",
        "GetSituacionActualJSon?Fecha=2026-09-28",
        "GetInyeccionesFisicasJSon?Year=2025",
    ]
    for i, c in enumerate(calls):
        r = get(WS + c, save=f"oc_ws_{i}.txt")
        if r is not None:
            print("   ", r.text[:1500].replace("\n", " "), flush=True)
    # Bring2mind DMX document folders (historic reports)
    for i, page in enumerate([
            "https://www.oc.org.do/Informes-Hist%C3%B3ricos/Operaci%C3%B3n-del-SENI/An%C3%A1lisis-Operativo",
            "https://www.oc.org.do/Informes-Hist%C3%B3ricos/Operaci%C3%B3n-del-SENI/Coordinaci%C3%B3n-y-Supervisi%C3%B3n-Tiempo-Real",
            "https://www.oc.org.do/Informes-Hist%C3%B3ricos/Operaci%C3%B3n-del-SENI/Programaci%C3%B3n-del-SENI",
            "https://www.oc.org.do/Informes/Operaci%C3%B3n-del-SENI/An%C3%A1lisis-Operativo",
            "https://www.oc.org.do/Informes/Operaci%C3%B3n-del-SENI/Coordinaci%C3%B3n-y-Supervisi%C3%B3n-Tiempo-Real"]):
        r = get(page, save=f"oc_page_{i}.html")
        if r is None:
            continue
        t = r.text
        for m in re.finditer(r"(Bring2mind[^\"'<>]{0,200}|moduleId[\"']?\s*[:=]\s*\d+|ModuleId=\d+|tabId[\"']?\s*[:=]\s*\d+|"
                             r"EntryId=\d+|Folder[A-Za-z]*[\"']?\s*[:=]\s*[\"']?\d+)", t):
            print("   ", m.group(0)[:200])
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", t, re.I | re.S):
            txt = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if re.search(r"diari|informe|operaci|generac|\.xls|\.pdf|Download", m.group(1) + txt, re.I):
                print("   A", urljoin(r.url, m.group(1))[:200], "|", txt[:80])


def mem():
    for i, page in enumerate([
            "https://mem.gob.do/category/sector-electrico/boletin-de-generacion-y-gestion-de-energia/",
            "https://mem.gob.do/category/sector-electrico/boletin-de-generacion-y-gestion-de-energia/2023-boletin-de-generacion/",
            "https://mem.gob.do/category/sector-electrico/informe-de-desempeno/",
            "https://mem.gob.do/category/estadisticas/",
            "https://mem.gob.do/category/sector-electrico/",
            "https://mem.gob.do/category/hidrocarburos/",
            "https://www.cne.gob.do/estadisticas/",
            "https://www.cne.gob.do/",
            "https://micm.gob.do/direcciones/combustibles/",
            "https://micm.gob.do/transparencia/estadisticas-institucionales/"]):
        r = get(page, save=f"mem_page_{i}.html")
        if r is None:
            continue
        seen = set()
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", r.text, re.I | re.S):
            link = urljoin(r.url, m.group(1))
            txt = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if link in seen:
                continue
            seen.add(link)
            if re.search(r"\.(xlsx?|pdf|ods|csv)|boletin|estad|desempe|gas|import|combust|categor", link + txt, re.I):
                print("   A", link[:220], "|", txt[:80])


if __name__ == "__main__":
    {"oc2": oc2, "mem": mem}[sys.argv[1]]()
    print("DONE", flush=True)
