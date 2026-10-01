"""
One-off probe: where Peru (COES), Ecuador (CENACE / CELEC) and Uruguay
(ADME / UTE / CTM Salto Grande) publish daily hydro reservoir levels or
volumes in a form a script can download, and how far back they go.

Every response is saved under hydro3_raw/ (the workflow pushes that to the
throwaway branch hydro3-discovery-raw for offline inspection) and logged
with status, type, size and keyword hits.

TLS: verification stays on everywhere. Servers that omit their
intermediate certificate get it added from the leaf's AIA URL (chain must
still end at a certifi root). For CELEC SUR's port-8443 API the leaf/issuer
details are printed so the right fix can be chosen.

Round 1 (ROUND=1): COES portal pages + scripts + file browser roots;
CENACE page keywords; CELEC SUR dashboard bundle + ORDS API; ADME per-plant
series columns and Salto Grande; UTE / CTM landing pages.
Round 1 findings:
  - COES portalinformacion has only 'generacion' and 'demanda' (no hydrology
    page); the file browser has 'Post Operacion/Reportes/IEOD/<year>/' folders
    (Informe de Evaluacion de la Operacion Diaria).
  - CELEC SUR's ORDS service (port 8443) serves a valid Sectigo chain, so
    plain certifi verification works. pointValuesMesH24 for Mazar (mrid
    30031) answers for 2021 too (not only from 2022) but leaves each month's
    last day empty. The dashboard bundle also calls pointValuesAnioH24 /
    AniosH24 / MesAvg / AnioAvg / AniosAvg. Plants: Mazar (cota 30031),
    Molino/Amaluza (24019), Sopladora (90919), Minas San Francisco (650919).
    Other CELEC units (Hidronacion: Daule-Peripa, Hidroagoyan: Pisayambo)
    have no such dashboard.
  - ADME seriescentralhidro.cgi: ids bon/bay/pal only (no Salto Grande);
    columns Pot_MW_, hToma_m_ (lake level), hDescarga_m_, QVertido, CE,
    QTurbinado, QErogado; Bonete has real values in 2019, none in 2015.
    adme.com.uy's menu links 'Operacion Rio Negro' = pronos.adme.com.uy/seriesbonete.php.
  - CTM Salto Grande publishes a daily 'Reporte de Caudales y Niveles' PDF
    (current day only) on datos_hidrologicos.php.
Round 2 (ROUND=2): COES IEOD folders down to one day's files (+ sitemap);
CELEC SUR year endpoints, history depth, month-end days; ADME
seriesbonete.php, Bonete history start; INA a5 Salto Grande level; CTM pages.
Round 2 findings:
  - COES IEOD day folders hold the daily annexes: 2026 'AnexoA_DDMM.xlsx'
    (sheet PRINCIP_VOLUMENES: half-hourly volumes/levels of ~29 reservoirs,
    mostly daily-regulation ponds; seasonal ones = Junin (Statkraft),
    Sibinacocha, Aricota, Macusani, Viconga, filled only at a few hours),
    2021 'Anexo2_Hidrologia_DDMM.xlsx' (sheet 'Princip_Caudales y Volumenes',
    hourly, 90 columns; Junin empty). Electroperu's Mantaro lagoons appear
    only as discharges (m3/s). No national useful-storage total in the IEOD.
  - CELEC SUR: Mazar level goes back to 2010 (pointValuesAnioH24 /
    MesH24); MesH24 with fechaFin = next month + 1 day returns every day
    including the last. Amaluza has 2021. 30032/30033 are other Mazar points
    (not levels).
  - ADME: Rio Negro levels (hToma) valid from 2017 (2016 = -121111 missing).
  - INA a5 series 26319 'Salto Grande Arriba' (Prefectura gauge, daily mean
    level, m) is current (34.38 m on 30-Sep-2026).
Round 3 (ROUND=3, Peru only): newest files under IDCOS, weekly / monthly
evaluation reports, daily and weekly operation programmes, bulletins; IEOD
hydrology annex layouts 2022-2025.
Round 3 findings: the weekly 'Informe Semanal de Evaluacion de la Operacion'
(Post Operacion/Informes/Evaluacion Semanal/<year>/SEMANAL N° w (dd.mm.yyyy -
dd.mm.yyyy)/Informe_Semanal_SEMw_yyyy.xlsx) has section 5.1 'Volumen util de
los embalses y lagunas': 29 reservoir/lagoon rows with start/end-of-week
useful volume, % and capacity (total 1,970 hm3), this year and last; 5.2
has 12 basin series by week for 4 years. IDCOS daily annex 9 has only a
few seasonal reservoirs. -> PERU_HYDRO_RESERVOIRS.py reads 5.1.
Round 4 (ROUND=4): the 5.1 test run failed on 2021-2023 reports (no sheet
with that title) - dump the volume sheets of older reports.
Round 4 findings: reports up to 2023 have one 'HIDROLOGIA' sheet of
company blocks (weekly by year since 1997: Yuracmayo, Edegel lagoons,
Electroperu sub-basins, Junin, EGASA, Viconga, Orazul, San Gaban, Aricota,
EGEMSA, Bamputane, Chalhuanca, Paucarcocha) that miss several reservoirs
and carry no capacities. From 2024 the 'Evolucion de volumenes' sheet (6.2
early 2024, 5.2 later) restates four years by basin (12 series; Junin is split
out of Mantaro between weeks 5 and 20 of 2024), so the last 2024 report gives
2021-2024. Capacities in table 5.1 change between reports (Junin 376 ->
315 hm3, San Gaban lagoons 376 -> 69 hm3). -> PERU_HYDRO_RESERVOIRS.py reads
the basin series of the newest report of each year from 2024.
"""

print("STARTING", flush=True)

import hashlib
import io
import json
import os
import re
import socket
import ssl
import sys
import time
from datetime import date, datetime, timedelta
from urllib.parse import quote, urljoin

import requests

OUT = "hydro3_raw"
os.makedirs(OUT, exist_ok=True)
ROUND = int(os.environ.get("ROUND", "1"))
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
      "Safari/537.36")
S = requests.Session()
S.headers["User-Agent"] = UA
KEY = re.compile(r"volumen|embalse|laguna|lago|hidrolog|cota|nivel|mazar|amaluza|daule|pisayambo|bonete|"
                 r"salto grande|junin|sibinacocha|hm3|mrid|ords|energ[ií]a embalsada|almacenad", re.I)
SAVED = []


def fname(name):
    return os.path.join(OUT, re.sub(r"[^A-Za-z0-9_.-]+", "_", name)[:150])


def get(url, name=None, method="GET", show=True, **kw):
    kw.setdefault("timeout", 90)
    for i in range(3):
        try:
            r = S.request(method, url, **kw)
            break
        except requests.RequestException as e:
            print(f"== {method} {url}: ERROR {type(e).__name__}: {str(e)[:300]}", flush=True)
            if i == 2:
                return None
            time.sleep(5 * (i + 1))
    ct = r.headers.get("Content-Type", "")
    print(f"== {method} {r.url}: {r.status_code} {ct} {len(r.content):,} bytes", flush=True)
    if name:
        with open(fname(name), "wb") as f:
            f.write(r.content)
        SAVED.append(name)
    if show and ("html" in ct or "json" in ct or "javascript" in ct or "text" in ct):
        hits(r.text)
    return r


def hits(text, width=140, limit=25):
    n = 0
    for m in KEY.finditer(text):
        a = max(0, m.start() - width // 2)
        snippet = re.sub(r"\s+", " ", text[a:m.end() + width // 2])
        print(f"   ~ {snippet}", flush=True)
        n += 1
        if n >= limit:
            print("   ~ (more hits cut)", flush=True)
            break


def links(html, base, pat=None):
    out = []
    for m in re.finditer(r"""(?:href|src)\s*=\s*["']([^"'#]+)["']""", html, re.I):
        u = urljoin(base, m.group(1))
        if (pat is None or re.search(pat, u, re.I)) and u not in out:
            out.append(u)
    return out


def cert_info(host, port):
    try:
        pem = ssl.get_server_certificate((host, port), timeout=30)
    except Exception as e:  # noqa: BLE001
        print(f"   cert {host}:{port}: ERROR {type(e).__name__}: {e}", flush=True)
        return None
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes
    c = x509.load_pem_x509_certificate(pem.encode())
    try:
        san = c.extensions.get_extension_for_class(x509.SubjectAlternativeName).value.get_values_for_type(x509.DNSName)
    except x509.ExtensionNotFound:
        san = []
    try:
        aia = str(c.extensions.get_extension_for_class(x509.AuthorityInformationAccess).value)
    except x509.ExtensionNotFound:
        aia = "-"
    print(f"   cert {host}:{port}: subject={c.subject.rfc4514_string()} issuer={c.issuer.rfc4514_string()} "
          f"san={san} notBefore={c.not_valid_before} notAfter={c.not_valid_after} "
          f"sha256={c.fingerprint(hashes.SHA256()).hex()} aia={aia[:300]}", flush=True)
    with open(fname(f"cert_{host}_{port}.pem"), "w") as f:
        f.write(pem)
    # the chain the server actually sends
    try:
        import subprocess
        out = subprocess.run(["openssl", "s_client", "-connect", f"{host}:{port}", "-servername", host, "-showcerts"],
                             input=b"", capture_output=True, timeout=40).stdout.decode("latin-1")
        with open(fname(f"chain_{host}_{port}.txt"), "w") as f:
            f.write(out)
        for line in out.splitlines():
            if re.match(r"\s*\d+ s:|\s+i:|Verify return code|Verification", line):
                print("   " + line.strip(), flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"   openssl: {e}", flush=True)
    return pem


def aia_bundle(host, port=443):
    """certifi roots + intermediates named in the leaf's AIA field (verification stays on)."""
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, port), timeout=30)
    cert = x509.load_pem_x509_certificate(pem.encode())
    extra = []
    for _ in range(3):
        try:
            aia = cert.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
        except x509.ExtensionNotFound:
            break
        urls = [d.access_location.value for d in aia if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS]
        if not urls:
            break
        der = requests.get(urls[0], timeout=30).content
        certs = (pkcs7.load_der_pkcs7_certificates(der) if urls[0].endswith(".p7c")
                 else [x509.load_der_x509_certificate(der)])
        extra += [c.public_bytes(Encoding.PEM).decode() for c in certs]
        cert = certs[0]
        if cert.issuer == cert.subject:
            break
    path = os.path.join(os.environ.get("RUNNER_TEMP", "/tmp"), f"{host}_{port}_ca_bundle.pem")
    with open(path, "w") as f:
        f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    print(f"   AIA bundle for {host}:{port}: {len(extra)} extra cert(s)", flush=True)
    return path


# ======================================================================
# PERU - COES
# ======================================================================
COES = "https://www.coes.org.pe/Portal/"


def coes_browse(path):
    r = get(COES + "browser/vistadatos", name=f"coes_browse_{path}", method="POST", show=False,
            data={"baseDirectory": path, "url": path, "indicador": "", "initialLink": "", "orderFolder": ""})
    if r is None or not r.ok:
        return []
    items = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (m.group(1), m.group(2))
        if it not in items:
            items.append(it)
    print(f"   browse '{path}': {len(items)} items: {items[:40]}", flush=True)
    return items


def peru():
    print("\n################ PERU (COES) ################", flush=True)
    pages = ["portalinformacion", "portalinformacion/generacion", "portalinformacion/hidrologia",
             "portalinformacion/Hidrologia", "portalinformacion/volumen", "portalinformacion/volumenes",
             "portalinformacion/embalses", "portalinformacion/caudales", "portalinformacion/recursoshidricos",
             "Hidrologia", "hidrologia", "PostOperacion/Reportes/Ieod", "PostOperacion/Reportes/Hidrologia",
             "Operacion/Hidrologia", "PostOperacion/Hidrologia"]
    scripts = []
    menu = set()
    for p in pages:
        r = get(COES + p, name=f"coes_{p}.html")
        if r is None or not r.ok:
            continue
        for u in links(r.text, r.url):
            if re.search(r"\.js(\?|$)", u) and "coes.org.pe" in u and not re.search(r"jquery|bootstrap|highcharts|"
                                                                                    r"moment|kendo|select2", u, re.I):
                if u not in scripts:
                    scripts.append(u)
            elif "coes.org.pe/Portal/" in u and not re.search(r"\.(css|png|jpg|gif|ico|svg|woff)", u, re.I):
                menu.add(u)
    print("\n-- COES links seen (filtered by keyword):", flush=True)
    for u in sorted(menu):
        if re.search(r"hidro|volum|embal|lagun|caudal|portalinformacion|ieod|reporte", u, re.I):
            print("   ", u, flush=True)
    print("\n-- COES scripts:", flush=True)
    for u in scripts[:40]:
        r = get(u, name="coesjs_" + u.split("/Portal/")[-1], show=False)
        if r is not None and r.ok:
            for m in re.finditer(r"""(?:url\s*:\s*|\$\.(?:post|get|ajax)\(\s*|controlador\s*\+\s*)["']([^"']+)["']""",
                                 r.text):
                print(f"   {u.split('/')[-1]}: {m.group(1)}", flush=True)
            hits(r.text, limit=8)
    # guessed JSON calls on the hydrology page, same form fields as the generation page
    today = date.today()
    d1, d0 = (today - timedelta(days=1)).strftime("%d/%m/%Y"), (today - timedelta(days=8)).strftime("%d/%m/%Y")
    for p in ["portalinformacion/hidrologia", "portalinformacion/Hidrologia"]:
        for data in [{"fechaInicial": d0, "fechaFinal": d1, "indicador": 0},
                     {"fechaInicial": d0, "fechaFinal": d1}]:
            r = get(COES + p, name=f"coes_post_{p}_{len(data)}.json", method="POST", data=data,
                    headers={"X-Requested-With": "XMLHttpRequest"})
            if r is not None and r.ok and "json" in r.headers.get("Content-Type", ""):
                try:
                    j = r.json()
                    print("   JSON keys:", list(j)[:30] if isinstance(j, dict) else type(j), flush=True)
                except ValueError:
                    pass
    # file browser: roots and IEOD
    for path in ["", "Post Operación/", "Post Operacion/", "Post Operación/Reportes/", "Post Operación/Reportes/IEOD/",
                 "Operación/", "Publicaciones/"]:
        coes_browse(path)


# ======================================================================
# ECUADOR - CENACE, CELEC SUR (Paute cascade), other CELEC units
# ======================================================================
def ecuador():
    print("\n################ ECUADOR ################", flush=True)
    cert_info("www.cenace.gob.ec", 443)
    try:
        bundle = aia_bundle("www.cenace.gob.ec")
        r = get("https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm", name="cenace_infoop.html",
                verify=bundle, timeout=120)
    except Exception as e:  # noqa: BLE001
        print("   CENACE failed:", type(e).__name__, e, flush=True)

    print("\n-- CELEC SUR dashboard", flush=True)
    host = "generacioncsr.celec.gob.ec"
    cert_info(host, 443)
    cert_info(host, 8443)
    r = get(f"https://{host}/graficasproduccion/", name="csr_dashboard.html")
    if r is not None and r.ok:
        for u in links(r.text, r.url, r"\.js(\?|$)"):
            js = get(u, name="csrjs_" + u.split("/")[-1], show=False)
            if js is None or not js.ok:
                continue
            t = js.text
            for m in re.finditer(r"https?://[^\s\"'`]+", t):
                if "celec" in m.group(0) or "ords" in m.group(0):
                    print("   URL in js:", m.group(0)[:200], flush=True)
            for m in re.finditer(r"mrid", t, re.I):
                print("   mrid ctx:", re.sub(r"\s+", " ", t[max(0, m.start() - 200):m.end() + 200]), flush=True)
            for m in re.finditer(r"Cota|Volumen|Nivel|Mazar|Amaluza|Sopladora|Embalse", t):
                print("   kw ctx:", re.sub(r"\s+", " ", t[max(0, m.start() - 150):m.end() + 150]), flush=True)
    base = f"https://{host}:8443/ords/csr/sardomcsr/"
    params = lambda mrid, a, b: {"mrid": mrid, "fechaInicio": f"{a}T00:00:00.000Z",  # noqa: E731
                                 "fechaFin": f"{b}T00:00:00.000Z", "fecha": f"{a[8:10]}/{a[5:7]}/{a[:4]} 00:00:00"}
    for verify_name, verify in [("certifi", True)]:
        r = get(base + "pointValuesMesH24", name="csr_mazar_2026-09.json", params=params(30031, "2026-09-01", "2026-10-01"),
                verify=verify)
    try:
        b = aia_bundle(host, 8443)
        r = get(base + "pointValuesMesH24", name="csr_mazar_2026-09_aia.json",
                params=params(30031, "2026-09-01", "2026-10-01"), verify=b)
        if r is not None and r.ok:
            for mon in ["2021-01-01", "2021-06-01", "2021-12-01", "2022-01-01"]:
                y, m = int(mon[:4]), int(mon[5:7])
                nxt = f"{y + (m == 12):04d}-{m % 12 + 1:02d}-01"
                get(base + "pointValuesMesH24", name=f"csr_mazar_{mon}.json", params=params(30031, mon, nxt), verify=b)
            # ORDS metadata / catalogue
            for p in ["", "metadata-catalog/", "open-api-catalog/", "points", "puntos"]:
                get(base + p, name=f"csr_meta_{p or 'root'}", verify=b)
    except Exception as e:  # noqa: BLE001
        print("   CELEC SUR 8443 AIA attempt failed:", type(e).__name__, e, flush=True)

    print("\n-- other CELEC units / INAMHI", flush=True)
    for u in ["https://www.celec.gob.ec/hidronacion/", "https://www.celec.gob.ec/hidroagoyan/",
              "https://www.celec.gob.ec/celecsur/", "https://www.celec.gob.ec/hidropaute/",
              "https://www.celec.gob.ec/", "https://generacionhan.celec.gob.ec/", "https://generacionhag.celec.gob.ec/",
              "https://www.inamhi.gob.ec/", "https://www.controlrecursosyenergia.gob.ec/estadisticas/"]:
        h = re.match(r"https://([^/]+)", u).group(1)
        try:
            r = get(u, name="ec_" + u.split("//")[1])
        except Exception as e:  # noqa: BLE001
            print("   ", u, type(e).__name__, e, flush=True)
            r = None
        if r is not None and r.ok:
            for l in links(r.text, r.url, r"cota|nivel|embalse|hidrolog|graficas|produccion|operaci|boletin|estad"):
                print("    link:", l, flush=True)


# ======================================================================
# URUGUAY - ADME, UTE, CTM Salto Grande
# ======================================================================
def uruguay():
    print("\n################ URUGUAY ################", flush=True)
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                                    "south_america"))
    import URUGUAY_ADME as U
    ser = lambda d: (d - date(1899, 12, 30)).days  # noqa: E731
    a, b = date.today() - timedelta(days=4), date.today()
    for plant in ["bon", "bay", "pal", "sg", "SG", "salto", "sgr", "ter", "terra"]:
        r = get(U.PLANT_SERIES_URL, name=f"adme_series_{plant}.ods",
                params={"idCentral": plant, "ts": "ods", "dtIni": ser(a), "dtFin": ser(b)}, show=False, timeout=180)
        if r is None or not r.ok:
            continue
        try:
            rows = next(iter(U.read_ods(r.content).values()))
            for row in rows[:6]:
                print(f"   {plant}: {row[:30]}", flush=True)
            print(f"   {plant}: {len(rows)} rows; last {rows[-1][:30] if rows else None}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"   {plant}: not ods ({type(e).__name__}); head {r.content[:300]!r}", flush=True)
    # how far back do the Rio Negro series go?
    for y in [2000, 2010, 2015, 2019]:
        r = get(U.PLANT_SERIES_URL, name=f"adme_series_bon_{y}.ods",
                params={"idCentral": "bon", "ts": "ods", "dtIni": ser(date(y, 1, 1)), "dtFin": ser(date(y, 1, 4))},
                show=False, timeout=180)
        if r is not None and r.ok:
            try:
                rows = next(iter(U.read_ods(r.content).values()))
                print(f"   bon {y}: {len(rows)} rows; sample {rows[min(5, len(rows) - 1)][:12]}", flush=True)
            except Exception as e:  # noqa: BLE001
                print(f"   bon {y}: {type(e).__name__}", flush=True)
    r = get("https://pronos.adme.com.uy/", name="adme_pronos_index.html")
    pages = set()
    if r is not None and r.ok:
        for u in links(r.text, r.url):
            if "adme.com.uy" in u and not re.search(r"\.(css|png|jpg|gif|ico|svg|js)(\?|$)", u, re.I):
                pages.add(u)
    for u in sorted(pages):
        print("    pronos link:", u, flush=True)
    for u in sorted(pages)[:60]:
        if re.search(r"cota|embal|hidro|serie|lago|bonete|central|estado|almac|energ", u, re.I):
            get(u, name="adme_" + u.split("//")[1])
    for u in ["https://pronos.adme.com.uy/cotas.php", "https://pronos.adme.com.uy/embalses.php",
              "https://pronos.adme.com.uy/hidro.php", "https://adme.com.uy/",
              "https://www.ute.com.uy/", "https://portal.ute.com.uy/", "https://www.saltogrande.org/",
              "https://www.saltogrande.org/es/", "https://www.saltogrande.org/hidrologia"]:
        r = get(u, name="uy_" + u.split("//")[1])
        if r is not None and r.ok:
            for l in links(r.text, r.url, r"cota|nivel|embalse|hidrolog|lago|represa|caudal|datos|estad"):
                print("    link:", l, flush=True)


# ---------------------------------------------------------------- round 2
def peru2():
    print("\n################ PERU round 2 ################", flush=True)
    r = get(COES + "sitemap/index", name="coes_sitemap.html", show=False)
    if r is not None and r.ok:
        for u in sorted(set(links(r.text, r.url, r"/Portal/"))):
            if not re.search(r"\.(css|png|jpg|gif|ico|svg|js)(\?|$)", u, re.I):
                print("   sitemap:", u, flush=True)
    get(COES + "PostOperacion/Reportes/Idcos", name="coes_idcos.html", show=False)
    for year in ["2026", "2021"]:
        months = [html_unescape(p) for p, k in coes_browse(f"Post Operación/Reportes/IEOD/{year}/") if k == "D"]
        months = sorted(set(months))
        if not months:
            continue
        mpath = months[-1] if year == "2026" else months[0]
        listing = coes_browse(mpath)
        days = sorted({html_unescape(p) for p, k in listing if k == "D"})
        files = [html_unescape(p) for p, k in listing if k == "F"]
        print(f"   {year}: month {mpath}: {len(days)} day folders {days[:3]}..., files {files[:10]}", flush=True)
        items = [(p, "F") for p in files]
        if days:
            dpath = days[-2] if len(days) > 1 else days[0]
            for p, k in coes_browse(dpath):
                p = html_unescape(p)
                items.append((p, k))
                if k == "D":
                    sub = coes_browse(p)
                    print(f"     sub {p}: {sub[:20]}", flush=True)
                    items += [(html_unescape(q), kk) for q, kk in sub if kk == "F"]
        seen = set()
        for p, k in items:
            if k != "F" or p in seen or not re.search(r"\.(xlsx?|xlsm|zip|csv)$", p, re.I):
                continue
            seen.add(p)
            r = get(COES + "browser/download?url=" + quote(p), name=f"ieod_{year}_" + p.split("/")[-1],
                    show=False, timeout=180)
            if r is None or not r.ok or not p.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue
            try:
                import pandas as pd
                xl = pd.ExcelFile(io.BytesIO(r.content))
                print(f"     sheets {p.split('/')[-1]}: {xl.sheet_names}", flush=True)
                for sh in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name=sh, header=None)
                    txt = df.astype(str).apply(lambda c: " | ".join(c), axis=1)
                    hit = txt[txt.str.contains("VOLUMEN|EMBALSE|LAGUNA|JUNIN|SIBINACOCHA|HIDROLOG", case=False)]
                    if len(hit):
                        print(f"       [{sh}] {len(df)}x{df.shape[1]} hits:", flush=True)
                        for i, t in hit.head(12).items():
                            print(f"         r{i}: {t[:300]}", flush=True)
            except Exception as e:  # noqa: BLE001
                print("     excel read failed:", type(e).__name__, e, flush=True)


def html_unescape(s):
    import html
    return html.unescape(s)


def ecuador2():
    print("\n################ ECUADOR round 2 ################", flush=True)
    base = "https://generacioncsr.celec.gob.ec:8443/ords/csr/sardomcsr/"

    def call(ep, mrid, a, b, name):
        r = get(base + ep, name=name, show=False, params={"mrid": mrid, "fechaInicio": f"{a}T00:00:00.000Z",
                                                          "fechaFin": f"{b}T00:00:00.000Z",
                                                          "fecha": f"{a[8:10]}/{a[5:7]}/{a[:4]} 00:00:00"})
        if r is None or not r.ok:
            return []
        try:
            j = r.json()
        except ValueError:
            print("   not json", r.text[:200], flush=True)
            return []
        it = j.get("items") or []
        vals = [x for x in it if x.get("valueedit") is not None]
        meta = {k: v for k, v in j.items() if k not in ("items", "links")}
        print(f"   {ep} {mrid} {a}..{b}: {len(it)} items, {len(vals)} with values; {meta}; "
              f"first {it[-1] if it else None}; last {it[0] if it else None}", flush=True)
        return it
    for y in [2026, 2021, 2018, 2015, 2012, 2010]:
        call("pointValuesAnioH24", 30031, f"{y}-01-01", f"{y + 1}-01-01", f"csr_anioh24_mazar_{y}.json")
    call("pointValuesAnioH24", 24019, "2021-01-01", "2022-01-01", "csr_anioh24_amaluza_2021.json")
    call("pointValuesAniosH24", 30031, "2010-01-01", "2026-10-01", "csr_aniosh24_mazar.json")
    for y in [2019, 2016, 2013]:
        call("pointValuesMesH24", 30031, f"{y}-03-01", f"{y}-04-01", f"csr_mesh24_mazar_{y}03.json")
    call("pointValuesMesH24", 30031, "2026-08-01", "2026-09-02", "csr_mesh24_mazar_aug_plus1.json")
    call("pointValuesMesH24", 30031, "2026-08-31", "2026-09-01", "csr_mesh24_mazar_aug31.json")
    call("pointValues", 30031, "2026-08-31", "2026-09-01", "csr_hourly_mazar_aug31.json")
    call("pointValues", 30031, "2026-09-30", "2026-10-01", "csr_hourly_mazar_sep30.json")
    call("pointValuesMesAvg", 30031, "2026-08-01", "2026-09-01", "csr_mesavg_mazar_aug.json")
    for mrid in [30030, 30032, 30033, 30039, 30040, 24018, 24020]:
        call("pointValuesMesH24", mrid, "2026-08-01", "2026-09-01", f"csr_probe_{mrid}.json")


def uruguay2():
    print("\n################ URUGUAY round 2 ################", flush=True)
    r = get("https://pronos.adme.com.uy/seriesbonete.php", name="adme_seriesbonete.html")
    if r is not None and r.ok:
        for u in links(r.text, r.url):
            if not re.search(r"\.(css|png|jpg|gif|ico|svg)(\?|$)", u, re.I):
                print("    link:", u, flush=True)
        for m in re.finditer(r"<form.*?</form>", r.text, re.S | re.I):
            print("    form:", re.sub(r"\s+", " ", m.group(0))[:800], flush=True)
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                                    "south_america"))
    import URUGUAY_ADME as U
    ser = lambda d: (d - date(1899, 12, 30)).days  # noqa: E731
    for y in [2016, 2017, 2018]:
        for plant in ["bon", "pal", "bay"]:
            r = get(U.PLANT_SERIES_URL, params={"idCentral": plant, "ts": "ods", "dtIni": ser(date(y, 6, 1)),
                                                "dtFin": ser(date(y, 6, 3))}, show=False, timeout=180)
            if r is not None and r.ok:
                try:
                    rows = next(iter(U.read_ods(r.content).values()))
                    print(f"   {plant} {y}: {rows[10][:4] if len(rows) > 10 else rows[-1]}", flush=True)
                except Exception as e:  # noqa: BLE001
                    print(f"   {plant} {y}: {type(e).__name__}", flush=True)
    r = get("https://alerta.ina.gob.ar/a5/obs/puntual/series/26319/observaciones", name="ina_26319.json", show=False,
            params={"timestart": "2026-09-01", "timeend": "2026-10-02", "format": "json"}, timeout=180)
    if r is not None and r.ok:
        obs = r.json()
        print(f"   INA 26319: {len(obs)} obs; {obs[:2]} .. {obs[-1:]}", flush=True)
    r = get("https://alerta.ina.gob.ar/a5/obs/puntual/series/26319", name="ina_26319_meta.json", show=False)
    if r is not None and r.ok:
        print("   INA 26319 meta:", r.text[:800], flush=True)
    for u in ["https://www.saltogrande.org/datos_hidrologicos.php", "https://www.saltogrande.org/datos_operativos.php"]:
        r = get(u, name="ctm_" + u.split("/")[-1])
        if r is not None and r.ok:
            for l in links(r.text, r.url):
                if re.search(r"docs|pdf|xls|csv|php", l, re.I):
                    print("    link:", l, flush=True)
    r = get("https://www.saltogrande.org/docs/hidrologia/CaudalesNiveles.pdf", name="ctm_CaudalesNiveles.pdf",
            show=False)
    if r is not None and r.ok:
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                for pg in pdf.pages[:2]:
                    print("    pdf:", (pg.extract_text() or "")[:1500], flush=True)
        except Exception as e:  # noqa: BLE001
            print("    pdf failed", e, flush=True)


# ---------------------------------------------------------------- round 3
HIT = re.compile(r"VOLUMEN|VOL\. ?UTIL|EMBALSE|LAGUNA|JUNIN|JUNÍN|SIBINACOCHA|HIDROLOG|ALMACEN|RESERVA H", re.I)


def scan_excel(content, label, rows=14):
    """Print sheet names and the rows that mention volumes / lagoons (strings only)."""
    import pandas as pd
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print(f"     {label}: not an Excel file we can read ({type(e).__name__}: {e})", flush=True)
        return
    print(f"     sheets {label.split('/')[-1]}: {xl.sheet_names}", flush=True)
    for sh in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sh, header=None)
        except Exception as e:  # noqa: BLE001
            print(f"       [{sh}] read failed {e}", flush=True)
            continue
        txt = df.apply(lambda r: " | ".join(str(v) for v in r if pd.notna(v)), axis=1)
        hit = txt[txt.str.contains(HIT)]
        if len(hit):
            print(f"       [{sh}] {df.shape} hits:", flush=True)
            for i, t in hit.head(rows).items():
                print(f"         r{i}: {t[:400]}", flush=True)
                # and the row below (often the values)
                if i + 1 in txt.index:
                    print(f"         r{i + 1}: {txt[i + 1][:400]}", flush=True)


def scan_zip(content, label):
    import zipfile
    try:
        z = zipfile.ZipFile(io.BytesIO(content))
    except zipfile.BadZipFile:
        print(f"     {label}: bad zip", flush=True)
        return
    names = z.namelist()
    print(f"     zip {label.split('/')[-1]}: {names[:30]}", flush=True)
    for n in names:
        if re.search(r"\.(xlsx?|xlsm)$", n, re.I) and HIT.search(n + " hidro volumen"):
            scan_excel(z.read(n), f"{label}!{n}", rows=10)


def coes_base_paths(page):
    r = get(COES + page, name="coes_" + page, show=False)
    if r is None or not r.ok:
        return []
    import html
    t = html.unescape(r.text)
    paths = sorted(set(re.findall(r"""["']((?:Post Operación|Operación|Publicaciones|Planificación)/[^"'<>]{2,120}/)["']""", t)))
    print(f"   {page}: base paths {paths}", flush=True)
    return paths


def newest_leaf(path, depth=4, want_files=2):
    """Walk down the newest folder until files appear; return (folder, files)."""
    for _ in range(depth):
        items = [(html_unescape(p), k) for p, k in coes_browse(path)]
        files = [p for p, k in items if k == "F"]
        dirs = [p for p, k in items if k == "D"]
        if len(files) >= want_files or not dirs:
            return path, files
        # newest = sort by name, prefer year/month/day numbering
        path = sorted(dirs, key=lambda d: re.sub(r"\D", "", d.rstrip("/").split("/")[-1]).zfill(8) + d)[-1]
    items = [(html_unescape(p), k) for p, k in coes_browse(path)]
    return path, [p for p, k in items if k == "F"]


def peru3():
    print("\n################ PERU round 3 ################", flush=True)
    pages = ["PostOperacion/Reportes/Idcos", "PostOperacion/Informes/EvaluacionSemanal",
             "PostOperacion/Informes/EvaluacionMensual", "Operacion/ProgOperacion/ProgramaDiario",
             "Operacion/ProgOperacion/ProgSemanalOp", "Publicaciones/Boletines/", "Publicaciones/Informes/"]
    downloads = 0
    for page in pages:
        for base in coes_base_paths(page)[:2]:
            folder, files = newest_leaf(base)
            print(f"   {page}: newest folder {folder}: {files[:25]}", flush=True)
            for f in files:
                if downloads >= 18:
                    break
                if not re.search(r"\.(xlsx?|xlsm|zip)$", f, re.I):
                    continue
                if re.search(r"cmg|costo|mantto|manten|rpf|rsf|hop", f, re.I):
                    continue
                r = get(COES + "browser/download?url=" + quote(f), name="pe3_" + f.split("/")[-1], show=False,
                        timeout=240)
                downloads += 1
                if r is None or not r.ok:
                    continue
                if f.lower().endswith(".zip"):
                    scan_zip(r.content, f)
                else:
                    scan_excel(r.content, f)
    # older IEOD layouts: one day per year 2022-2025 (which file holds the volumes, and is Junin filled?)
    for y in ["2022", "2023", "2024", "2025"]:
        months = sorted({html_unescape(p) for p, k in coes_browse(f"Post Operación/Reportes/IEOD/{y}/") if k == "D"})
        if not months:
            continue
        days = sorted({html_unescape(p) for p, k in coes_browse(months[len(months) // 2]) if k == "D"})
        if not days:
            continue
        files = [html_unescape(p) for p, k in coes_browse(days[len(days) // 2]) if k == "F"]
        print(f"   IEOD {y}: {days[len(days) // 2]}: {[f.split('/')[-1] for f in files]}", flush=True)
        for f in files:
            if re.search(r"hidro|anexoa|anexo2", f, re.I) and f.lower().endswith((".xlsx", ".xls")):
                r = get(COES + "browser/download?url=" + quote(f), name="pe3_ieod_" + f.split("/")[-1], show=False,
                        timeout=240)
                if r is not None and r.ok:
                    scan_excel(r.content, f, rows=6)


def peru4():
    """Older weekly reports: which sheet holds the useful-volume table before 2024, and in what layout."""
    print("\n################ PERU round 4 ################", flush=True)
    picks = {"2020": [10, 40], "2021": [43], "2022": [28], "2023": [12, 48], "2024": [5, 20]}
    for y, wks in picks.items():
        items = [(html_unescape(p), k) for p, k in coes_browse(f"Post Operación/Informes/Evaluacion Semanal/{y}/")]
        for p, k in items:
            m = re.search(r"N\D{0,3}(\d+)\s*\(", p)
            if k != "D" or not m or int(m.group(1)) not in wks:
                continue
            files = [(html_unescape(q), kk) for q, kk in coes_browse(p)]
            print(f"   {p}: {files}", flush=True)
            for f, kk in files:
                if kk != "F" or not re.search(r"\.(xlsx?|xlsm|zip)$", f, re.I):
                    continue
                r = get(COES + "browser/download?url=" + quote(f), name=f"pe4_{y}_" + f.split("/")[-1], show=False,
                        timeout=240)
                if r is None or not r.ok:
                    continue
                if f.lower().endswith(".zip"):
                    scan_zip(r.content, f)
                    continue
                import pandas as pd
                try:
                    xl = pd.ExcelFile(io.BytesIO(r.content))
                except Exception as e:  # noqa: BLE001
                    print("     unreadable", type(e).__name__, e, flush=True)
                    continue
                print(f"     sheets: {xl.sheet_names}", flush=True)
                for sh in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name=sh, header=None, nrows=60)
                    txt = " | ".join(str(v) for v in df.values.ravel() if isinstance(v, str))
                    if re.search(r"VOL[UÚ]MEN", txt, re.I) and re.search(r"EMBALSE|LAGUNA", txt, re.I):
                        print(f"     [{sh}] {df.shape}", flush=True)
                        for i, row in df.iterrows():
                            cells = [str(v).replace(chr(10), " ")[:45] for v in row if pd.notna(v)]
                            if cells:
                                print(f"        r{i}: {cells[:12]}", flush=True)


ROUND_FUNCS = {2: [("pe", peru2), ("ec", ecuador2), ("uy", uruguay2)], 3: [("pe", peru3)], 4: [("pe", peru4)]}


if __name__ == "__main__":
    which = os.environ.get("ONLY", "pe,ec,uy").split(",")
    for k, fn in ROUND_FUNCS.get(ROUND, [("pe", peru), ("ec", ecuador), ("uy", uruguay)]):
        if k in which:
            try:
                fn()
            except Exception as e:  # noqa: BLE001
                import traceback
                traceback.print_exc()
    print(f"\nSaved {len(SAVED)} raw files in {OUT}/", flush=True)
