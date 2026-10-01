"""
Round 10 (Oct-2026), Honduras. NORTH_CENTRAL_AMERICA_POWER_PROBE9.py found
the ODS app page 'produccion-horaria?p8_indx=<technology>' = interactive
report of hourly MW per plant (Fecha, Planta, 00..23) for a date range
P8_FECHA_INICIAL..P8_FECHA_FINAL (m/d/yyyy), earliest P8_MIN_DATE
(6/1/2026 on 30-Sep-2026: a rolling ~4-month window). Monthly
'Informe Mensual Operacion del Mercado' PDFs carry 'Tabla 8 Produccion
General de Energia (GWh)'.
Here: technology index -> name (from the 'otr' donut series), a CSV
download for a date range per index, and the Tabla 8 text of one monthly
and one annual report.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE10.py ods
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


def ctx(t):
    g = lambda p: (re.search(p, t).group(1) if re.search(p, t) else "")  # noqa: E731
    return {"sess": g(r'name="p_instance" value="(\d+)"'), "step": g(r'name="p_flow_step_id" value="(\d+)"'),
            "flow": g(r'name="p_flow_id" value="(\d+)"'), "salt": g(r'value="(\d+)" id="pSalt"'),
            "prot": g(r'id="pPageItemsProtected" value="([^"]+)"').replace("&#x2F;", "/")}


def ods():
    S.verify = aia_bundle("cnd.enee.hn")
    t = S.get(APP + "otr", timeout=60).text
    c = ctx(t)
    print("otr ctx", {k: v for k, v in c.items() if k != "prot"}, flush=True)
    m = re.search(r'jetChart\.init\("R56798483702719832".*?"(UkVHSU9O[^"]+)"\)', t, re.S)
    if m:
        data = {"p_flow_id": c["flow"], "p_flow_step_id": c["step"], "p_instance": c["sess"], "p_debug": "",
                "p_request": "PLUGIN=" + m.group(1).encode().decode("unicode_escape"),
                "p_json": json.dumps({"pageItems": {"itemsToSubmit": [], "protected": c["prot"], "rowVersion": "",
                                                    "formRegionChecksums": []}, "salt": c["salt"]})}
        j = S.post(BASE + "wwv_flow.ajax", data=data, timeout=60).json()
        for s in j.get("series", []):
            it = s["items"][0]
            print(f"  DONUT {s['name'].strip()!r}: {it['value']} MWh today, link {it.get('link', '')[-60:]}", flush=True)
    sess = c["sess"]
    # each technology index: first plants for 1 day, and a CSV for a 3-day range
    for indx in range(1, 13):
        url = f"{APP}producci%C3%B3n-horaria?p8_indx={indx}&session={sess}"
        r = S.get(url, timeout=60)
        text = html.unescape(re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<[^>]+>", " ", r.text, flags=re.S)))
        plants = re.findall(r"\d+/\d+/\d{4} ([A-Za-zÁÉÍÓÚÑáéíóúñ][^\d]{1,40}?) [\d.]", text)
        pc = ctx(r.text)
        print(f"  INDX {indx}: [{r.status_code}] page={pc['step']} plants={sorted(set(plants))[:14]}", flush=True)
        if indx in (1, 6):
            for q in [f"f?p={pc['flow']}:{pc['step']}:{pc['sess']}:CSV:::P8_INDX,P8_FECHA_INICIAL,P8_FECHA_FINAL:{indx},6/1/2026,6/3/2026",
                      f"f?p={pc['flow']}:{pc['step']}:{pc['sess']}:CSV:::P8_INDX,P8_FECHA_INICIAL,P8_FECHA_FINAL:{indx},5/1/2026,5/2/2026"]:
                cr = S.get(BASE + q, timeout=60)
                body = cr.content.decode("latin-1", "replace")
                lines = body.splitlines()
                print(f"   CSV {q[-60:]}: [{cr.status_code}] {cr.headers.get('content-type')} {len(lines)} lines", flush=True)
                for ln in lines[:4] + lines[-2:]:
                    print("     ", ln[:200], flush=True)
    # Tabla 8 of the newest monthly report and of an annual report
    import pdfplumber
    for pid, want in [(10, 0), (41, 2)]:
        t = S.get(f"{BASE}f?p=110:4:::::p4_id:{pid}", timeout=60).text
        g = lambda p: (re.search(p, t).group(1) if re.search(p, t) else "")  # noqa: E731
        data = {"p_flow_id": "110", "p_flow_step_id": "4", "p_instance": g(r'name="p_instance" value="(\d+)"'),
                "p_debug": "", "p_request": "PLUGIN=" + g(r'"ajaxIdentifier":"([^"]+)"').encode().decode("unicode_escape"),
                "p_widget_name": "worksheet", "p_widget_mod": "ACTION", "p_widget_action": "LAZY_LOAD",
                "p_widget_num_return": "250", "x01": g(r'_worksheet_id" value="(\d+)"'), "x02": g(r'_report_id" value="(\d+)"'),
                "p_json": json.dumps({"pageItems": {"itemsToSubmit": [{"n": "P4_ID", "v": str(pid), "ck": g(r'data-for="P4_ID" value="([^"]+)"')}],
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
        print(f"\n  LIST {pid}: {[c[0] for c, _ in rows]}", flush=True)
        if len(rows) <= want:
            continue
        b = S.get(BASE + rows[want][1][0], timeout=300).content
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            print(f"  == {rows[want][0][0]}: {len(pdf.pages)} pages", flush=True)
            for i, p in enumerate(pdf.pages[:40]):
                txt = p.extract_text() or ""
                if re.search(r"Producci[oó]n General de Energ|seg[uú]n su tipo de Tecnolog|PRODUCCI[OÓ]N DE ENERG[IÍ]A BRUTA", txt, re.I) \
                        and re.search(r"\d,\d{3}|\d\.\d{2}", txt):
                    print(f"  --- page {i + 1}\n{txt[:3500]}", flush=True)


if __name__ == "__main__":
    {"ods": ods}[sys.argv[1]]()
    print("DONE", flush=True)
