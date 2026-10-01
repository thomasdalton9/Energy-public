"""
Round 4 (Oct-2026) after NORTH_CENTRAL_AMERICA_POWER_PROBE3.py found:
  AMM  graph.js calls /graficaTipoRecurso?dt=<datearea> and
       /graficaCombustible?dt=<datearea> (datearea is dd/mm/yyyy);
       GM<date>.xlsx has a 'GENERACION POR TIPO DE RECURSO' monthly block.
  UT   ut.com.sv's own name servers (190.242.x / 190.120.15.100) don't answer
       outside El Salvador; Cloudflare holds a stale www -> 190.120.15.116.
  ODS  the CND reports are Oracle APEX pages under
       https://appcnd.enee.hn:3200/odsprd/f?p=110:<page>:::::p<page>_id:<n>
       (informe diario = 110:6 id 2, informes diarios OM = 110:4 id 6,
       operacion mensual = 110:4 id 504), plus a real-time 'otr' page.
  EOR  only MER transactions; the SCADA map service is Cloudflare-blocked.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE4.py amm|ut|ods
"""

print("STARTING", flush=True)

import datetime as dt
import re
import ssl
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9"}
S = requests.Session()
S.headers.update(H)


def get(url, **kw):
    try:
        r = S.get(url, timeout=60, **kw)
        print(f"[{r.status_code}] {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:300]}", flush=True)
        return None


def amm():
    r = get("https://wl12.amm.org.gt/GraficaPW/js/graph.js")
    if r is not None:
        lits = sorted(set(re.findall(r"'(/[A-Za-z][^'\s]{2,60})'", r.text)))
        print("  literals:", lits, flush=True)
        m = re.search(r"dirURLBase\s*=\s*([^;,]+)", r.text)
        print("  dirURLBase:", m.group(0)[:200] if m else None, flush=True)
    base = "https://wl12.amm.org.gt/GraficaPW"
    today = dt.date(2026, 10, 1)
    dates = [today - dt.timedelta(days=1), today - dt.timedelta(days=6), dt.date(2026, 9, 1), dt.date(2026, 6, 15),
             dt.date(2025, 6, 15), dt.date(2023, 3, 10), dt.date(2021, 1, 1)]
    for ep in ["/graficaTipoRecurso", "/graficaCombustible"]:
        for d in dates:
            for fmt in ["%d/%m/%Y"]:
                r = get(base + ep, params={"dt": d.strftime(fmt)})
                if r is not None and r.ok:
                    t = r.text
                    print(f"    {ep} {d}: {len(t)} chars: {t[:700]}", flush=True)
    r = get(base + "/graficaTipoRecurso", params={"dt": "1/9/2026"})
    if r is not None:
        print("    non-padded:", r.text[:200], flush=True)


def ut():
    import urllib3.util.connection as uc
    orig = uc.create_connection
    pins = {"www.ut.com.sv": "190.120.15.116", "ut.com.sv": "190.120.15.116",
            "estadistico.ut.com.sv": "190.120.15.116"}

    def patched(address, *a, **kw):
        host, port = address
        return orig((pins.get(host, host), port), *a, **kw)
    uc.create_connection = patched  # name resolution only; TLS still verifies the real host name
    for u in ["https://www.ut.com.sv/", "https://estadistico.ut.com.sv/OperacionDiaria.aspx", "http://www.ut.com.sv/"]:
        r = get(u)
        if r is not None:
            print("   ", re.sub(r"\s+", " ", r.text[:1500]), flush=True)
    uc.create_connection = orig


def aia_bundle(host, port=443):
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, port))  # read the leaf only to find its AIA URL
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
        certs = pkcs7.load_der_pkcs7_certificates(der) if urls[0].endswith(".p7c") else [x509.load_der_x509_certificate(der)]
        extra += [c.public_bytes(Encoding.PEM).decode() for c in certs]
        cert = certs[0]
        if cert.issuer == cert.subject:
            break
    path = f"/tmp/{host}_bundle.pem"
    open(path, "w").write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    return path


def apex_dump(url, label):
    r = get(url)
    if r is None:
        return None
    t = r.text
    title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
    print(f"  == {label}: title={title.group(1).strip()[:80] if title else ''}", flush=True)
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", t, re.I | re.S):
        h, tx = m.group(1), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
        if re.search(r"apex_util|download|\.pdf|\.xls|f\?p=|r/|javascript:apex", h, re.I) and "p_lang" not in h:
            print("    A", urljoin(r.url, h)[:250], "|", tx[:80], flush=True)
    for m in re.finditer(r"<(select|input)[^>]*(id|name)=\"(P\d+_[A-Z0-9_]+)\"[^>]*>", t, re.I):
        print("    ITEM", m.group(0)[:200], flush=True)
    text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", t, flags=re.S))
    print("    TEXT", text[:1800], flush=True)
    for m in re.finditer(r"ajaxIdentifier\"?\s*:\s*\"([^\"]+)\"", t):
        print("    AJAX", m.group(1)[:120], flush=True)
    return r


def ods():
    S.verify = aia_bundle("cnd.enee.hn")
    try:
        b2 = aia_bundle("appcnd.enee.hn", 3200)
        open(S.verify, "a").write("\n" + open(b2).read().split(open(__import__("certifi").where()).read())[-1])
    except Exception as e:  # noqa: BLE001
        print("appcnd bundle:", e, flush=True)
    base = "https://appcnd.enee.hn:3200/odsprd/"
    for label, p in [("informe diario", "f?p=110:6:::::p6_id:2"), ("informes diarios OM", "f?p=110:4:::::p4_id:6"),
                     ("operacion mensual", "f?p=110:4:::::p4_id:504"), ("operacion mercado", "f?p=110:4:::::p4_id:10"),
                     ("anual mercado", "f?p=110:4:::::p4_id:41"), ("trimestral", "f?p=110:4:::::p4_id:21"),
                     ("otr", "ods_prd/r/operador-del-sistema-ods/otr")]:
        apex_dump(base + p, label)
    for u in ["https://cnd.enee.hn/2022/11/18/hidroelectrica/", "https://cnd.enee.hn/2023/01/20/energia_generada-png/"]:
        r = get(u)
        if r is not None:
            for m in re.finditer(r"<(img|iframe)[^>]*src=[\"']([^\"']+)", r.text, re.I):
                if "wp-content/uploads" in m.group(2) or "appcnd" in m.group(2):
                    print("   ", m.group(2)[:200], flush=True)


if __name__ == "__main__":
    {"amm": amm, "ut": ut, "ods": ods}[sys.argv[1]]()
    print("DONE", flush=True)
