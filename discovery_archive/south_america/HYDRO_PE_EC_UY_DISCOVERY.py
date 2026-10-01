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
"""

print("STARTING", flush=True)

import hashlib
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


if __name__ == "__main__":
    which = os.environ.get("ONLY", "pe,ec,uy").split(",")
    for k, fn in [("pe", peru), ("ec", ecuador), ("uy", uruguay)]:
        if k in which:
            try:
                fn()
            except Exception as e:  # noqa: BLE001
                import traceback
                traceback.print_exc()
    print(f"\nSaved {len(SAVED)} raw files in {OUT}/", flush=True)
