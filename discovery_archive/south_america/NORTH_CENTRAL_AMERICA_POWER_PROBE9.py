"""
Round 9 (Oct-2026), Honduras. NORTH_CENTRAL_AMERICA_POWER_PROBE8.py found:
  - the daily PDF's 'produccion bruta por tipo' is a Power BI donut whose
    values cannot be tied to labels from the text layer;
  - the real-time page 'otr' has JET charts with 1-minute MW by technology
    for the last 24 h and today's energy by technology; each point links to
    page 'produccion-horaria?p8_indx=<tech>'.
  - monthly 'Informe Mensual Operacion del Mercado' PDFs (p4_id 10, from
    Mar-2024) and annual reports (p4_id 41, 2019-2025).
Here: what 'produccion-horaria' (and the app's other pages) offer - a date
item would give history; and the tables in one monthly report.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE9.py ods
"""

print("STARTING", flush=True)

import html
import io
import json
import re
import ssl
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9"}
S = requests.Session()
S.headers.update(H)
BASE = "https://appcnd.enee.hn:3200/odsprd/"
APP = BASE + "r/ods_prd/operador-del-sistema-ods/"


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


def dump(url, label):
    r = S.get(url, timeout=60)
    t = r.text
    title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S)
    print(f"\n== {label}: [{r.status_code}] {r.url[:160]} title={title.group(1).strip() if title else ''} {len(t):,} chars", flush=True)
    for m in re.finditer(r"<(input|select)[^>]*(?:id|name)=\"(P\d+_[A-Z0-9_]+)\"[^>]*>", t):
        print("   ITEM", re.sub(r"\s+", " ", m.group(0))[:220], flush=True)
    for m in re.finditer(r"<select[^>]*id=\"(P\d+_[A-Z0-9_]+)\"[^>]*>(.*?)</select>", t, re.S):
        opts = re.findall(r"<option[^>]*value=\"([^\"]*)\"[^>]*>(.*?)</option>", m.group(2))
        print(f"   SELECT {m.group(1)}: {opts[:20]}", flush=True)
    for m in re.finditer(r'(jetChart\.init\("[^"]+"|interactiveReport\(\{"regionId":"[^"]+"|interactiveGrid|"componentName":"[^"]+")', t):
        print("   REGION", m.group(1)[:160], flush=True)
    for m in re.finditer(r'<a\b[^>]*href="([^"]+)"[^>]*>(.*?)</a>', t, re.S):
        h = html.unescape(m.group(1))
        if "operador-del-sistema-ods/" in h and "login" not in h:
            print("   NAV", h[:160], "|", re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:60], flush=True)
    text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", t, flags=re.S))
    print("   TEXT", text[:1200], flush=True)
    return t


def ods():
    S.verify = aia_bundle("cnd.enee.hn")
    t = dump(APP + "otr", "otr")
    sess = re.search(r"session=(\d+)", t)
    sess = sess.group(1) if sess else ""
    for page in ["producci%C3%B3n-horaria?p8_indx=1", "par%C3%A1metros-generaci%C3%B3n", "par%C3%A1metros-del-sistema1",
                 "reserva", "home"]:
        dump(APP + page + ("&" if "?" in page else "?") + f"session={sess}", page)
    # one monthly market report: pages with generation-by-type tables
    import pdfplumber
    r = S.get(f"{BASE}f?p=110:4:::::p4_id:10", timeout=60)
    t = r.text
    g = lambda p: (re.search(p, t).group(1) if re.search(p, t) else "")  # noqa: E731
    data = {"p_flow_id": "110", "p_flow_step_id": "4", "p_instance": g(r'name="p_instance" value="(\d+)"'),
            "p_debug": "", "p_request": "PLUGIN=" + g(r'"ajaxIdentifier":"([^"]+)"').encode().decode("unicode_escape"),
            "p_widget_name": "worksheet", "p_widget_mod": "ACTION", "p_widget_action": "LAZY_LOAD",
            "p_widget_num_return": "250", "x01": g(r'_worksheet_id" value="(\d+)"'), "x02": g(r'_report_id" value="(\d+)"'),
            "p_json": json.dumps({"pageItems": {"itemsToSubmit": [{"n": "P4_ID", "v": "10", "ck": g(r'data-for="P4_ID" value="([^"]+)"')}],
                                                "protected": g(r'id="pPageItemsProtected" value="([^"]+)"').replace("&#x2F;", "/"),
                                                "rowVersion": "", "formRegionChecksums": []},
                                  "salt": g(r'value="(\d+)" id="pSalt"')})}
    htm = S.post(BASE + "wwv_flow.ajax", data=data, timeout=60).text
    rows = []
    for m in re.finditer(r"<tr[^>]*>(.*?)</tr>", htm, re.S):
        cells = [html.unescape(re.sub(r"<[^>]+>|\s+", " ", x)).strip() for x in re.findall(r"<td[^>]*>(.*?)</td>", m.group(1), re.S)]
        links = [html.unescape(h) for h in re.findall(r'href="([^"]+)"', m.group(1)) if "get_blob" in h]
        if cells:
            rows.append((cells, links))
    print("\nMONTHLY REPORTS:", [c[0] for c, _ in rows], flush=True)
    if rows:
        b = S.get(BASE + rows[0][1][0], timeout=180).content
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            for i, p in enumerate(pdf.pages[:16]):
                txt = p.extract_text() or ""
                if re.search(r"PRODUCCI[OÓ]N|Tabla|GWh|MWh", txt, re.I):
                    print(f"  --- page {i + 1}\n{txt[:2200]}", flush=True)


if __name__ == "__main__":
    {"ods": ods}[sys.argv[1]]()
    print("DONE", flush=True)
