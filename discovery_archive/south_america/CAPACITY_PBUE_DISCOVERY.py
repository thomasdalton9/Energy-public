"""
Discovery round 1: installed / effective generation capacity by technology
for Peru (COES, MINEM), Bolivia (CNDC), Uruguay (ADME, MIEM, UTE) and
Ecuador (ARCERNNR, CENACE). Prints status, size and keyword-matching links
for each seed page, the CNDC wp-json route index and document types, and
Ember's yearly capacity rows for validation. TLS verification stays on; a
host that omits its intermediate certificate is retried with certifi + the
intermediate named in the leaf's AIA field.

Usage: python3 CAPACITY_PBUE_DISCOVERY.py [section ...]   (peru bolivia uruguay ecuador ember)
"""
import io
import json
import re
import ssl
import sys
from urllib.parse import urljoin, urlparse

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"}
T = (15, 60)
KEY = re.compile(r"capacid|potencia|instalad|efectiv|estad|anuario|boletin|bolet%C3%ADn|informe|parque|"
                 r"\.xlsx?\b|\.pdf\b|\.csv\b|\.zip\b|balance|generaci|centrales", re.I)
BUNDLES = {}


def aia_bundle(host, port=443):
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, port), timeout=20)  # read only, to find the AIA URL
    cert = x509.load_pem_x509_certificate(pem.encode())
    try:
        san = cert.extensions.get_extension_for_oid(ExtensionOID.SUBJECT_ALTERNATIVE_NAME).value
        print(f"    leaf subject={cert.subject.rfc4514_string()} SAN={san.get_values_for_type(x509.DNSName)[:10]} "
              f"notAfter={cert.not_valid_after}")
    except Exception as e:  # noqa: BLE001
        print(f"    leaf subject={cert.subject.rfc4514_string()} ({e})")
    extra = []
    for _ in range(3):
        try:
            aia = cert.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
        except x509.ExtensionNotFound:
            break
        urls = [d.access_location.value for d in aia if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS]
        if not urls:
            break
        print(f"    AIA: {urls[0]}")
        der = requests.get(urls[0], timeout=30).content
        certs = (pkcs7.load_der_pkcs7_certificates(der) if urls[0].endswith(".p7c")
                 else [x509.load_der_x509_certificate(der)])
        extra += [c.public_bytes(Encoding.PEM).decode() for c in certs]
        cert = certs[0]
        if cert.issuer == cert.subject:
            break
    path = f"/tmp/{host}_bundle.pem"
    with open(path, "w") as f:
        f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    return path


def get(url, method="GET", quiet=False, **kw):
    host = urlparse(url).hostname
    verify = BUNDLES.get(host, True)
    for attempt in range(2):
        try:
            r = requests.request(method, url, headers=H, timeout=T, verify=verify, **kw)
            if not quiet:
                print(f"  [{r.status_code}] {r.url} {r.headers.get('content-type', '')} {len(r.content)}B")
            return r
        except requests.exceptions.SSLError as e:
            print(f"  SSL {url}: {str(e)[:200]}")
            if attempt == 0:
                try:
                    verify = BUNDLES[host] = aia_bundle(host)
                    continue
                except Exception as e2:  # noqa: BLE001
                    print(f"    AIA retry failed: {type(e2).__name__}: {str(e2)[:200]}")
            return None
        except requests.RequestException as e:
            print(f"  ERR {url}: {type(e).__name__}: {str(e)[:200]}")
            return None


def links(url, pat=KEY, limit=80, show=True):
    r = get(url)
    if r is None or r.status_code != 200 or "html" not in r.headers.get("content-type", ""):
        return []
    found = []
    for href, text in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        full = urljoin(r.url, href.strip())
        if pat.search(full) or pat.search(text):
            if full not in [f for f, _ in found]:
                found.append((full, text[:80]))
    if show:
        for f, t in found[:limit]:
            print(f"     - {t!r} -> {f}")
        if len(found) > limit:
            print(f"     ... {len(found) - limit} more")
    return found


def text_snippets(url, pat, width=200, limit=15):
    r = get(url)
    if r is None or r.status_code != 200:
        return
    for m in list(re.finditer(pat, r.text, re.I))[:limit]:
        s = re.sub(r"\s+", " ", r.text[max(0, m.start() - width):m.end() + width])
        print(f"     ~ {s}")


def peru():
    print("\n================ PERU")
    for u in ["https://www.coes.org.pe/Portal/portalinformacion",
              "https://www.coes.org.pe/Portal/publicaciones/estadisticas/",
              "https://www.coes.org.pe/Portal/Publicaciones/Estadisticas/estadanual",
              "https://www.coes.org.pe/Portal/publicaciones/estadisticas/estadmensual",
              "https://www.coes.org.pe/Portal/publicaciones/boletinmensual",
              "https://www.coes.org.pe/Portal/",
              "https://www.minem.gob.pe/_estadistica.php?idSector=6",
              "https://www.gob.pe/institucion/minem/colecciones/16283-anuario-estadistico-de-electricidad"]:
        print(f"\n -- {u}")
        links(u)
    for u in ["https://www.coes.org.pe/Portal/portalinformacion/potenciainstalada",
              "https://www.coes.org.pe/Portal/portalinformacion/potenciaefectiva",
              "https://www.coes.org.pe/Portal/portalinformacion/capacidad",
              "https://www.coes.org.pe/Portal/portalinformacion/produccion",
              "https://www.coes.org.pe/Portal/portalinformacion/maximademanda"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            print("     ", re.sub(r"\s+", " ", r.text[:300]))


def bolivia():
    print("\n================ BOLIVIA")
    r = get("https://www.cndc.bo/wp-json/cndc/v1")
    if r is not None and r.status_code == 200:
        try:
            for route, info in r.json().get("routes", {}).items():
                eps = info.get("endpoints") or [{}]
                print(f"     route {route} {info.get('methods')} args={list(eps[0].get('args', {}).keys())}")
        except ValueError:
            print(r.text[:500])
    for cat in [155, None]:
        params = {"agrupado": "true"}
        if cat:
            params["categoria_id"] = cat
        r = get("https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos", params=params)
        if r is None or r.status_code != 200:
            continue
        try:
            j = r.json()
        except ValueError:
            print(r.text[:300])
            continue
        grupos = j.get("grupos", []) if isinstance(j, dict) else j
        print(f"     {len(grupos)} groups; keys={list(j.keys()) if isinstance(j, dict) else type(j)}")
        types = {}
        for g in grupos:
            for d in g.get("documentos", []) if isinstance(g, dict) else []:
                t = d.get("tipo_documento") or d.get("titulo") or str(d)[:60]
                types.setdefault(t, []).append((g.get("label") or g.get("periodo") or g.get("mes"),
                                                d.get("archivo_url")))
        if grupos:
            print("     first group:", json.dumps(grupos[0], ensure_ascii=False)[:800])
        for t, v in types.items():
            print(f"     type {t!r}: {len(v)} docs; latest {v[0]}; oldest {v[-1]}")
    for u in ["https://www.cndc.bo/wp-json/cndc/v1/estadisticas/categorias",
              "https://www.cndc.bo/wp-json/cndc/v1/historico/capacidad?desde=2015&hasta=2026",
              "https://www.cndc.bo/wp-json/cndc/v1/historico/potencia?desde=2015&hasta=2026",
              "https://www.cndc.bo/wp-json/cndc/v1/capacidad",
              "https://www.cndc.bo/wp-json/cndc/v1/agentes"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            print("     ", r.text[:600])
    print("\n -- cndc.bo pages")
    for u in ["https://www.cndc.bo/", "https://www.cndc.bo/estadisticas/", "https://www.cndc.bo/agentes/generacion/"]:
        links(u, limit=60)


def uruguay():
    print("\n================ URUGUAY")
    r = get("https://catalogodatos.gub.uy/api/3/action/package_search", params={"q": "potencia instalada", "rows": 20})
    if r is not None and r.status_code == 200:
        for p in r.json()["result"]["results"]:
            print(f"     pkg {p['name']}: {p['title']}")
            for res in p.get("resources", [])[:8]:
                print(f"         {res.get('format')} {res.get('name')} {res.get('url')}")
    for u in ["https://www.adme.com.uy/", "https://pronos.adme.com.uy/",
              "https://www.adme.com.uy/informes.php", "https://adme.com.uy/mmee/infanual.php",
              "https://ben.miem.gub.uy/", "https://ben.miem.gub.uy/descargas.html",
              "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos",
              "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/estadisticas",
              "https://portal.ute.com.uy/institucional/infraestructura/fuentes-de-generacion"]:
        print(f"\n -- {u}")
        links(u)
    text_snippets("https://portal.ute.com.uy/institucional/infraestructura/fuentes-de-generacion", r"MW")


def ecuador():
    print("\n================ ECUADOR")
    for u in ["https://www.controlrecursosyenergia.gob.ec/",
              "https://www.controlrecursosyenergia.gob.ec/estadistica-del-sector-electrico/",
              "https://www.controlrecursosyenergia.gob.ec/estadisticas-del-sector-electrico/",
              "https://www.controlrecursosyenergia.gob.ec/estadistica-del-sector-electrico-ecuatoriano/",
              "https://www.cenace.gob.ec/", "https://www.cenace.gob.ec/informe-anual/",
              "https://www.cenace.gob.ec/informes-anuales/",
              "https://www.recursosyenergia.gob.ec/", "https://www.datosabiertos.gob.ec/dataset/?q=potencia"]:
        print(f"\n -- {u}")
        links(u)


def ember():
    print("\n================ EMBER yearly capacity")
    import pandas as pd
    url = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/yearly_full_release_long_format.csv"
    r = get(url)
    if r is None or r.status_code != 200:
        return
    d = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    print("     columns:", list(d.columns))
    c = d[(d["Area"].isin(["Peru", "Bolivia", "Uruguay", "Ecuador"])) & (d["Category"] == "Capacity")]
    print("     units:", c["Unit"].unique(), "subcats:", c["Subcategory"].unique())
    c = c[(c["Subcategory"] == "Fuel") & (c["Year"] >= 2020)]
    print(c.pivot_table(index=["Area", "Year"], columns="Variable", values="Value", aggfunc="sum").round(3).to_string())


if __name__ == "__main__":
    which = sys.argv[1:] or ["peru", "bolivia", "uruguay", "ecuador", "ember"]
    for w in which:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
