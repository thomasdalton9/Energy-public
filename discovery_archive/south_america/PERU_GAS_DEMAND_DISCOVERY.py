"""
One-off probe: where can Peru natural gas demand by sector (power,
industrial, vehicle CNG, residential/commercial), production by lot and
Peru LNG exports be downloaded from, monthly from 2021?

Crawls the candidate publishers (MINEM, Osinergmin, Perupetro, TGP,
Calidda, COES, BCRP, datosabiertos.gob.pe), logs every link that looks
relevant and saves the pages and data files under pe_raw/ (the workflow
pushes them to a throwaway branch for offline inspection).
Round 1 findings: MINEM (gob.pe) collection 17643 'Informes Estadisticos
Upstream - Downstream' has a monthly 'distribucion-<mes>-<anio>.xlsx'
(Informe de distribucion de gas natural): one sheet per distribution
concession (Calidda, Contugas, Quavii, Petroperu, Gasnorp) with monthly
volume by sector in MMPCD for a rolling window (Jan-2022..latest).
Osinergmin's quarterly Boletin Estadistico de Gas Natural only has charts.
www.minem.gob.pe times out; perupetro.com.pe fails TLS verification
(incomplete chain); datosabiertos CKAN API 404.

Round 2 (ROUND=2): list every publication in collection 17643 (all pages)
and save their attachments (distribution, upstream) for offline inspection;
retry Perupetro with its missing intermediate CA added to the trust store
(fetched from the leaf's AIA 'CA Issuers' URL - verification stays on).
Round 2 findings: collection 17643 runs Jan-2023..Aug-2026 (88 monthly
publications). Downstream 'distribucion-<mes>-<anio>.xlsx' files were all
re-uploaded in 2026 and every one now covers Jan-2022..its month (no 2021).
Upstream '7-produccion-fiscalizada-de-gas-natural.xlsx' = monthly fiscalised
gas production by lot (MPCD) from Jan-2019. Perupetro works with the
intermediate CA added: its Estadisticas menu has Produccion Diaria, Gas
Natural, Embarques de GN para exportacion, Informe Mensual. Calidda links a
'Reportes Regulatorios' page (appadmin.calidda.com.pe/ReportesOsinerming).

Round 3 (ROUND=3): Perupetro statistics pages + their files, Calidda
regulatory reports, gob.pe searches for 2021-2022 distribution reports.
Round 3 findings: Perupetro statistics are PDFs (daily 'Reporte de
Produccion de Gas' per month, 'Estadistica Mensual', annual books); LNG
cargoes sit in an iframe /ExportaGAS/Relacion_ES.jsp. Calidda's
'Reportes Regulatorios' (ASP.NET WebForms, postbacks) lists daily
'Reporte Operativo Volumetrico', 'Reportes Operativos' (GNLC) and monthly
'Reporte Operativo por Categoria Tarifaria' - 2,856 files. gob.pe search
results are rendered client-side (no hits server-side).

Round 4 (ROUND=4): Calidda report lists per type (first and last page)
and downloads; Perupetro LNG cargo list and sample PDFs.
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


COLL = "https://www.gob.pe/institucion/minem/colecciones/17643-informes-estadisticos-upstream-downstream"


def round2():
    import subprocess
    import certifi
    os.makedirs(RAW, exist_ok=True)
    s = requests.Session()
    pubs = {}
    for page in range(1, 40):
        r = get(f"{COLL}?sheet={page}", s)
        if r is None or r.status_code != 200:
            break
        save(f"{COLL}_sheet{page}.html", r.content)
        found = re.findall(r'href="(/institucion/minem/informes-publicaciones/(\d+)-[^"?#]+)"', r.text)
        new = [u for u, _ in found if u not in pubs]
        for u in new:
            pubs[u] = None
        out(f"  sheet {page}: {len(new)} new publications")
        if not new:
            break
    out(f"{len(pubs)} publications in the collection")
    for u in pubs:
        out(f"PUB {u}")
    files = 0
    for u in pubs:
        if time.time() > DEADLINE:
            break
        r = get("https://www.gob.pe" + u, s)
        if r is None or r.status_code != 200:
            continue
        att = sorted(set(re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+', r.text)))
        for a in att:
            out(f"  ATT {a}")
        keep = [a for a in att if a.lower().endswith((".xlsx", ".xls", ".zip", ".csv"))
                and re.search(r"distrib|gas|upstream|produc|fiscaliz|liquid|regal|hidrocarb", a, re.I)
                and not re.search(r"refiner|inventar|precios-comb|ventas", a, re.I)]
        for a in keep:
            if files >= MAX_FILES:
                break
            rr = get(a, s)
            if rr is not None and rr.status_code == 200:
                files += 1
                out(f"  FILE saved {save(a, rr.content)}")
    # Perupetro: add the missing intermediate CA (from the leaf's AIA) to certifi's bundle; verification stays on
    host = "www.perupetro.com.pe"
    try:
        pem = subprocess.run(["openssl", "s_client", "-connect", f"{host}:443", "-servername", host, "-showcerts"],
                             input=b"", capture_output=True, timeout=30).stdout.decode("latin-1")
        out(pem[:3000])
        leaf = re.search(r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----", pem, re.S).group(0)
        txt = subprocess.run(["openssl", "x509", "-noout", "-text"], input=leaf.encode(), capture_output=True).stdout.decode()
        aia = re.findall(r"CA Issuers - URI:(\S+)", txt)
        out(f"AIA: {aia}")
        bundle = os.path.join(RAW, "perupetro_ca.pem")
        extra = ""
        for url in aia:
            der = requests.get(url, headers=H, timeout=T).content
            conv = subprocess.run(["openssl", "x509", "-inform", "DER" if not der.startswith(b"-----") else "PEM"],
                                  input=der, capture_output=True).stdout.decode()
            extra += conv
        with open(bundle, "w") as f:
            f.write(open(certifi.where()).read() + "\n" + extra)
        s.verify = bundle
        for u in ["https://www.perupetro.com.pe/", 
                  "https://www.perupetro.com.pe/wps/portal/corporativo/PerupetroSite/estadisticas",
                  "https://www.perupetro.com.pe/wps/portal/corporativo/PerupetroSite/estadisticas/estadistica%20petrolera"]:
            r = get(u, s)
            if r is not None and r.status_code == 200:
                save(u + ".html", r.content)
                for l in sorted(set(links(r.text, r.url))):
                    if re.search(r"estad|produc|gas|xls|pdf", l, re.I):
                        out(f"  LINK {l}")
    except Exception as e:  # noqa: BLE001
        out(f"perupetro probe failed: {type(e).__name__}: {e}")
    out(f"round 2 done: {files} files")


def perupetro_bundle():
    """certifi + Perupetro's missing intermediate CA (from the leaf's AIA URL); verification stays on."""
    import subprocess
    import certifi
    host = "www.perupetro.com.pe"
    pem = subprocess.run(["openssl", "s_client", "-connect", f"{host}:443", "-servername", host, "-showcerts"],
                         input=b"", capture_output=True, timeout=30).stdout.decode("latin-1")
    leaf = re.search(r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----", pem, re.S).group(0)
    txt = subprocess.run(["openssl", "x509", "-noout", "-text"], input=leaf.encode(), capture_output=True).stdout.decode()
    extra = ""
    for url in re.findall(r"CA Issuers - URI:(\S+)", txt):
        der = requests.get(url, headers=H, timeout=T).content
        extra += subprocess.run(["openssl", "x509", "-inform", "PEM" if der.startswith(b"-----") else "DER"],
                                input=der, capture_output=True).stdout.decode()
    bundle = os.path.join(RAW, "perupetro_ca.pem")
    with open(bundle, "w") as f:
        f.write(open(certifi.where()).read() + "\n" + extra)
    return bundle


PP = "https://www.perupetro.com.pe/wps/portal/corporativo/PerupetroSite/estadisticas/"
PP_PAGES = {"produccion_diaria": "Z6_N2E4HH41JGPM80Q4TDJTM80BM2", "gas_natural": "Z6_N2E4HH41JGPM80Q4TDJTM80FP3",
            "embarques_gn": "Z6_N2E4HH41JGPM80Q4TDJTM80HU6", "estadistica_petrolera": "Z6_N2E4HH41JGPM80Q4TDJTM801L4",
            "informe_mensual": "Z6_N2E4HH41JGPM80Q4TDJTM80JH1", "lgn": "Z6_N2E4HH41JGPM80Q4TDJTM80V47",
            "compromisos_l88": "Z6_N2E4HH41JGPM80Q4TDJTM80RJ0", "industria_gn": "Z6_N2E4HH41JGPM80Q4TDJTM80L30"}


def crawl_files(s, url, tag, max_files=12, follow=1):
    r = get(url, s)
    if r is None or r.status_code != 200:
        return
    save(f"{tag}.html", r.content)
    ls = list(dict.fromkeys(links(r.text, r.url)))
    files = [u for u in ls if re.search(r"\.(xlsx?|csv|zip|pdf)(\?|$)|MOD=AJPERES|/documentos?/|download", u, re.I)]
    for u in ls:
        if re.search(r"estad|produc|gas|embarq|mensual|report|volum|categor", u, re.I):
            out(f"  LINK {u[:220]}")
    n = 0
    for u in files:
        out(f"  FILELINK {u[:220]}")
        if n < max_files and not re.search(r"\.pdf(\?|$)", u, re.I):
            rr = get(u, s)
            if rr is not None and rr.status_code == 200:
                n += 1
                out(f"  FILE saved {save(tag + '_' + u.split('/')[-1][:80], rr.content)}")


def round3():
    os.makedirs(RAW, exist_ok=True)
    s = requests.Session()
    try:
        s.verify = perupetro_bundle()
    except Exception as e:  # noqa: BLE001
        out(f"bundle failed: {e}")
    for tag, oid in PP_PAGES.items():
        crawl_files(s, f"{PP}?uri=nm:oid:{oid}", "pp_" + tag)
    for u in ["https://appadmin.calidda.com.pe/ReportesOsinerming/Reportes/documentos",
              "https://appadmin.calidda.com.pe/ReportesOsinerming/",
              "https://www.calidda.com.pe/ir/es/"]:
        crawl_files(s, u, "cal_" + re.sub(r"\W+", "_", u.split("//")[1])[:60], max_files=20)
    for term in ["distribucion de gas natural 2021", "informe estadistico downstream 2022",
                 "volumen de gas natural distribuido por sector", "estadistica de gas natural 2021"]:
        u = "https://www.gob.pe/busquedas?term=" + requests.utils.quote(term) + "&institucion[]=minem"
        r = get(u, s)
        if r is not None and r.status_code == 200:
            save("search_" + term, r.content)
            for m in re.finditer(r'href="(/institucion/minem/informes-publicaciones/[^"]+)"', r.text):
                out(f"  SEARCHHIT {m.group(1)}")
    out("round 3 done")


CAL = "https://appadmin.calidda.com.pe/ReportesOsinerming/Reportes/documentos"


def html_unescape(x):
    import html
    return html.unescape(x)


def aspnet_fields(text):
    return {m.group(1): html_unescape(m.group(2))
            for m in re.finditer(r'<input type="hidden" name="([^"]+)" id="[^"]*" value="([^"]*)"', text)}


def grid_rows(text):
    rows = []
    for m in re.finditer(r'lblTipoReporte_(\d+)"[^>]*>([^<]*)</span>\s*</td><td>([^<]*)</td><td[^>]*>([^<]*)</td>', text):
        rows.append((int(m.group(1)), html_unescape(m.group(2)), html_unescape(m.group(3)), m.group(4)))
    return rows


def cal_post(s, text, extra):
    data = aspnet_fields(text)
    data.update({"ctl00$main$ddlTipoReporte": extra.pop("_tipo", ""),
                 "ctl00$main$ucwPaginacion$ddlTamanioGrilla": extra.pop("_size", "30")})
    if "_page" in extra:
        data["ctl00$main$ucwPaginacion$ddlPaginacion"] = extra.pop("_page")
    data.update(extra)
    return s.post(CAL, data=data, headers=H, timeout=T)


def cal_download(s, text, idx, tipo, page=None):
    extra = {"_tipo": tipo, f"ctl00$main$gdvArchivos$ctl{idx + 2:02d}$btnDescargar.x": "8",
             f"ctl00$main$gdvArchivos$ctl{idx + 2:02d}$btnDescargar.y": "8", "__EVENTTARGET": ""}
    if page:
        extra["_page"] = page
    r = cal_post(s, text, extra)
    cd = r.headers.get("content-disposition", "")
    out(f"  download idx {idx}: {r.status_code} {r.headers.get('content-type')} {len(r.content)}b cd={cd}")
    if r.status_code == 200 and "html" not in r.headers.get("content-type", ""):
        name = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', cd)
        out(f"  FILE saved {save('cal_' + tipo + '_' + (name.group(1) if name else str(idx)), r.content)}")


def round4():
    os.makedirs(RAW, exist_ok=True)
    s = requests.Session()
    r = get(CAL, s)
    for tipo in ["CategoriaTarifaria", "Volumetricos", "GNLC"]:
        try:
            r1 = cal_post(s, r.text, {"_tipo": tipo, "__EVENTTARGET": "ctl00$main$ddlTipoReporte"})
            r1 = cal_post(s, r1.text, {"_tipo": tipo, "__EVENTTARGET": "ctl00$main$ucwPaginacion$ddlTamanioGrilla"})
            save(f"cal_{tipo}_p1.html", r1.content)
            total = re.search(r"Registros encontrados:\s*(\d+)", r1.text)
            sel = re.findall(r'ddlPaginacion.*?</select>', r1.text, re.S)
            npages = len(re.findall(r"<option", sel[0])) if sel else 1
            out(f"== {tipo}: {total.group(1) if total else '?'} records, {npages} pages")
            for row in grid_rows(r1.text):
                out(f"  ROW {row}")
            for i in range(2):
                cal_download(s, r1.text, i, tipo)
            if npages > 1:
                rl = cal_post(s, r1.text, {"_tipo": tipo, "_page": str(npages),
                                           "__EVENTTARGET": "ctl00$main$ucwPaginacion$ddlPaginacion"})
                save(f"cal_{tipo}_plast.html", rl.content)
                rows = grid_rows(rl.text)
                for row in rows:
                    out(f"  LASTROW {row}")
                if rows and tipo == "CategoriaTarifaria":
                    cal_download(s, rl.text, rows[-1][0], tipo, page=str(npages))
                if tipo == "CategoriaTarifaria":
                    for pg in range(max(2, npages - 6), npages):
                        rp = cal_post(s, r1.text, {"_tipo": tipo, "_page": str(pg),
                                                   "__EVENTTARGET": "ctl00$main$ucwPaginacion$ddlPaginacion"})
                        for row in grid_rows(rp.text):
                            out(f"  P{pg} ROW {row}")
        except Exception as e:  # noqa: BLE001
            out(f"{tipo} failed: {type(e).__name__}: {e}")
    try:
        s.verify = perupetro_bundle()
    except Exception as e:  # noqa: BLE001
        out(f"bundle failed: {e}")
    rr = get("https://www.perupetro.com.pe/ExportaGAS/Relacion_ES.jsp", s)
    if rr is not None and rr.status_code == 200:
        save("pp_exportagas.html", rr.content)
        for link in dict.fromkeys(links(rr.text, rr.url)):
            out(f"  LINK {link}")
    base = "https://www.perupetro.com.pe/wps/wcm/connect/corporativo/"
    for path in ["7d03f6b6-3d69-417c-9cbc-2bc791c7dc85/Reporte+de+Producci%C3%B3n+de+Gas+31.8.2026.pdf?MOD=AJPERES",
                 "5dbf87aa-b716-4934-9425-713810ca701a/Estadistica+Mensual-Junio+2026..pdf?MOD=AJPERES",
                 "be17839d-d5ae-4b17-89a9-6e5730e157ff/ESTAD%C3%8DSTICA%2BANUAL%2BDE%2BHIDROCARBUROS%2B2021%2B.pdf"
                 "?MOD=AJPERES"]:
        rr = get(base + path, s)
        if rr is not None and rr.status_code == 200:
            out(f"  FILE saved {save('pp_' + path.split('/')[1][:60] + '.pdf', rr.content)}")
    out("round 4 done")


if __name__ == "__main__":
    {"2": round2, "3": round3, "4": round4}.get(ROUND, main)()
