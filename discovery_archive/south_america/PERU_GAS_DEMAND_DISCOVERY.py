"""
One-off probe: where can Peru natural gas demand by sector (power,
industrial, vehicle CNG, residential/commercial), production by lot and
Peru LNG exports be downloaded from, monthly from 2021?

Crawls the candidate publishers (MINEM, Osinergmin, Perupetro, TGP,
Calidda, COES, BCRP, datosabiertos.gob.pe), logs every link that looks
relevant and saves the pages and data files under pe_raw/ (the workflow
pushes them to a throwaway branch for offline inspection).
Runs in GitHub Actions only (sites are blocked from the editing sandbox).
"""
import os
import re
import time
from collections import deque
from urllib.parse import urljoin, urlparse

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-PE,es;q=0.9,en;q=0.8"}
T = (10, 60)
RAW = "pe_raw"
ROUND = os.environ.get("ROUND", "1")

SEEDS = [
    # MINEM
    "https://www.gob.pe/institucion/minem/informes-publicaciones",
    "https://www.gob.pe/institucion/minem/colecciones",
    "https://www.gob.pe/busquedas?term=estadistica%20hidrocarburos&institucion[]=minem",
    "https://www.minem.gob.pe/_estadistica.php?idSector=5",
    "https://www.minem.gob.pe/_estadisticaSector.php?idSector=5",
    "https://www.minem.gob.pe/_publicacion.php?idSector=5",
    "https://www.minem.gob.pe/",
    # Osinergmin
    "https://www.gob.pe/institucion/osinergmin/informes-publicaciones",
    "https://www.gob.pe/institucion/osinergmin/colecciones",
    "https://www.osinergmin.gob.pe/",
    "https://www.osinergmin.gob.pe/empresas/gas-natural",
    "https://gasnatural.osinergmin.gob.pe/",
    "https://observatorio.osinergmin.gob.pe/",
    "https://www.osinergmin.gob.pe/seccion/institucional/acerca_osinergmin/estudios_economicos/reportes-de-mercado",
    "https://www.osinergmin.gob.pe/seccion/institucional/regulacion-tarifaria/publicaciones/gas-natural",
    # Perupetro
    "https://www.perupetro.com.pe/wps/portal/corporativo/PerupetroSite/estadisticas",
    "https://www.perupetro.com.pe/wps/portal/corporativo/PerupetroSite/estadisticas/estadistica%20petrolera",
    "https://www.perupetro.com.pe/",
    # TGP / Calidda / Peru LNG
    "https://www.tgp.com.pe/",
    "https://www.tgp.com.pe/es/",
    "https://www.calidda.com.pe/",
    "https://perulng.com/",
    # COES (fuel consumption)
    "https://www.coes.org.pe/Portal/PostOperacion/Reportes/Ieod",
    "https://www.coes.org.pe/Portal/portalinformacion",
    # open data / BCRP
    "https://www.datosabiertos.gob.pe/api/3/action/package_search?q=gas%20natural&rows=100",
    "https://www.datosabiertos.gob.pe/search/type/dataset?query=gas%20natural",
    "https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/produccion-minera-e-hidrocarburos",
    "https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/resultados/PN01800AM/html",
]
FOLLOW = re.compile(r"gas|estad|hidrocarb|boletin|bolet%c3%adn|demanda|consumo|produc|publicac|reporte|informe|anuario|"
                    r"mensual|datos|ieod|combust|camisea|volumen|transport|indicador|observatorio|colecci", re.I)
DATA_EXT = re.compile(r"\.(xlsx?|xlsm|csv|zip|json)(\?|$)", re.I)
PDF_EXT = re.compile(r"\.pdf(\?|$)", re.I)
PDF_KEEP = re.compile(r"gas|hidrocarb|estad|boletin|demanda|camisea", re.I)
ALLOWED = ("gob.pe", "perupetro.com.pe", "tgp.com.pe", "calidda.com.pe", "perulng.com", "coes.org.pe",
           "bcrp.gob.pe", "datosabiertos.gob.pe")
MAX_PAGES = int(os.environ.get("MAX_PAGES", "260"))
MAX_FILES = int(os.environ.get("MAX_FILES", "120"))
DEADLINE = time.time() + int(os.environ.get("BUDGET_S", "1500"))


def out(*a):
    print(*a, flush=True)


def save(url, content):
    name = re.sub(r"[^\w.\-]+", "_", urlparse(url).netloc + urlparse(url).path + ("_" + urlparse(url).query if urlparse(url).query else ""))[-150:]
    path = os.path.join(RAW, name)
    with open(path, "wb") as f:
        f.write(content)
    return path


def get(url, s):
    try:
        r = s.get(url, headers=H, timeout=T, allow_redirects=True)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b final={r.url if r.url != url else '='}")
        return r
    except requests.RequestException as e:
        out(f"GET {url} -> ERROR {type(e).__name__}: {str(e)[:160]}")
        return None


def links(html, base):
    for m in re.finditer(r"""(?:href|src|data-url)\s*=\s*["']([^"'#]+)["']""", html, re.I):
        u = urljoin(base, m.group(1).strip())
        if u.startswith("http"):
            yield u
    for m in re.finditer(r"""https?://[^\s"'<>\\]+\.(?:xlsx?|csv|zip|pdf)""", html, re.I):
        yield m.group(0)


def main():
    os.makedirs(RAW, exist_ok=True)
    s = requests.Session()
    q = deque((u, 0) for u in SEEDS)
    seen, pages, files = set(), 0, 0
    while q and time.time() < DEADLINE:
        url, depth = q.popleft()
        if url in seen:
            continue
        seen.add(url)
        host = urlparse(url).netloc
        if not any(host.endswith(a) for a in ALLOWED):
            continue
        is_data, is_pdf = bool(DATA_EXT.search(url)), bool(PDF_EXT.search(url))
        if is_data or is_pdf:
            if files >= MAX_FILES:
                continue
            r = get(url, s)
            if r is not None and r.status_code == 200 and len(r.content) < 25e6:
                files += 1
                out(f"  FILE saved {save(url, r.content)}")
            continue
        if pages >= MAX_PAGES:
            continue
        r = get(url, s)
        pages += 1
        if r is None or r.status_code != 200:
            continue
        ctype = r.headers.get("content-type", "")
        if "html" not in ctype and "json" not in ctype and "text" not in ctype:
            files += 1
            out(f"  FILE saved {save(url, r.content)} ({ctype})")
            continue
        save(url + ".html", r.content)
        text = r.text
        title = re.search(r"<title[^>]*>(.*?)</title>", text, re.I | re.S)
        out(f"  title: {title.group(1).strip()[:120] if title else '-'}")
        for u in dict.fromkeys(links(text, r.url)):
            if u in seen:
                continue
            if DATA_EXT.search(u):
                out(f"  DATA {u}")
                q.appendleft((u, depth + 1))
            elif PDF_EXT.search(u):
                if PDF_KEEP.search(u):
                    out(f"  PDF {u}")
                    if depth < 2:
                        q.append((u, depth + 1))
            elif depth < 2 and FOLLOW.search(u) and any(urlparse(u).netloc.endswith(a) for a in ALLOWED):
                out(f"  LINK {u}")
                q.append((u, depth + 1))
    out(f"done: {pages} pages, {files} files, {len(q)} left in queue")


if __name__ == "__main__":
    main()
