"""
Bolivia gas production and exports - round 2.

Round 1 (BOLIVIA_GAS_EXPORTS_DISCOVERY.py) found:
  - INE H01 / H02: monthly gas production (total; by department) in
    million m3, 2001-2025 - no field split.
  - INE 'Exportaciones segun Pais de Destino y Producto por Ano y Mes,
    2017-2026': value (USD) and net weight (tonnes) by country and product,
    to 2026. Volumes in m3 are not given there.
  - YPFB 'Boletin informativo' PDFs are newsletters (no statistics);
    MHE 'Boletin Energetico' is quarterly.
This round:
  1. INE H01/H02: full notes and the last rows.
  2. INE exports by country and product: every gas row under BRASIL and
     ARGENTINA (value and weight), last 24 months, and the month headers.
  3. INE exports database page (physical quantity in m3?).
  4. MHE 'Estadisticas' page and the export pages of the quarterly bulletin.
  5. ANH 'Planificacion y Estadistica', YPFB Transporte.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import io
import re
import ssl
import tempfile
from urllib.parse import urljoin

import certifi
import pandas as pd
import pdfplumber
import requests
from cryptography import x509
from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:160]}")
        return None


def links(html, base):
    res = []
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        res.append((t, urljoin(base, h.replace("&amp;", "&").replace("&#47;", "/"))))
    return res


out("=================== 1. INE H01 / H02 notes and last rows")
for u in ("https://nube.ine.gob.bo/index.php/s/Zc56MHV1Zrt0fJY/download",
          "https://nube.ine.gob.bo/index.php/s/QQ1WOj4zNwoeFlx/download"):
    r = get(u)
    if r is None or r.status_code != 200:
        continue
    for name, df in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
        out(f"  sheet {name!r} {df.shape}")
        for row in df.tail(40).itertuples(index=False):
            vals = [str(x) for x in row if str(x) != "nan"]
            if vals:
                out("     ", vals)

out("\n=================== 2. INE exports by country and product: gas rows")
r = get("https://nube.ine.gob.bo/index.php/s/0ToRunCr9uUpLOH/download")
if r is not None and r.status_code == 200:
    for name, df in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
        out(f"\n  sheet {name!r} {df.shape}")
        hdr = [i for i in range(12) if str(df.iat[i, 0]).strip().upper().startswith("PA")]
        h = hdr[0] if hdr else 4
        years = df.iloc[h].tolist()
        months = df.iloc[h + 1].tolist()
        ncol = df.shape[1]
        last = [j for j in range(1, ncol) if df.iloc[h + 3:, j].notna().any()]
        lastc = max(last) if last else ncol - 1
        out(f"  header row {h}; last col with data {lastc}: {years[lastc]} {months[lastc]}")
        cols = list(range(max(1, lastc - 23), lastc + 1))
        out("  cols:", [f"{years[j]}-{months[j]}" for j in cols])
        country = None
        for i in range(h + 2, len(df)):
            lab = str(df.iat[i, 0]).strip()
            if lab and lab != "nan" and lab.upper() == lab and not re.search(r"\d", lab):
                country = lab
            if re.search(r"gas|GLP|petr", lab, re.I) or country in ("BRASIL", "ARGENTINA") and lab == country:
                out(f"   [{country}] {lab!r}: {[round(float(df.iat[i, j]), 3) if pd.notna(df.iat[i, j]) else None for j in cols]}")
        # 2021 Jan column onwards for gas rows to Brazil / Argentina
        start = [j for j in range(1, ncol) if str(years[j]).startswith("2021") and str(months[j]).strip().lower() == "enero"]
        if start:
            country = None
            for i in range(h + 2, len(df)):
                lab = str(df.iat[i, 0]).strip()
                if lab and lab != "nan" and lab.upper() == lab and not re.search(r"\d", lab):
                    country = lab
                if re.search(r"gas natural", lab, re.I):
                    vals = [round(float(df.iat[i, j]), 2) if pd.notna(df.iat[i, j]) else None
                            for j in range(start[0], lastc + 1)]
                    out(f"   2021-> [{country}] {lab!r}: {vals}")

out("\n=================== 3. INE exports database page")
r = get("https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/bases-de-datos-exportaciones/")
if r is not None and r.status_code == 200:
    for t, u in links(r.text, r.url):
        if "nube.ine" in u or re.search(r"\.(xlsx?|csv|zip|rar)$", u, re.I):
            out(f"   LINK {t[:110]!r} -> {u}")


def bundle_with_intermediates(host):
    """certifi roots + intermediates fetched via the leaf's AIA URLs (verification stays on)."""
    pem = ssl.get_server_certificate((host, 443), timeout=30)
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
        raw = requests.get(urls[0], headers=H, timeout=T).content
        certs = []
        for loader in (lambda b: [x509.load_der_x509_certificate(b)], lambda b: [x509.load_pem_x509_certificate(b)],
                       pkcs7.load_der_pkcs7_certificates, pkcs7.load_pem_pkcs7_certificates):
            try:
                certs = loader(raw)
                break
            except Exception:
                continue
        if not certs:
            break
        extra += [c.public_bytes(Encoding.PEM).decode() for c in certs]
        cert = certs[0]
        if cert.subject == cert.issuer:
            break
    f = tempfile.NamedTemporaryFile("w", suffix=".pem", delete=False)
    f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    f.close()
    return f.name


out("\n=================== 4. MHE")
try:
    bundle = bundle_with_intermediates("www.mhe.gob.bo")
    r = get("https://www.mhe.gob.bo/estadisticas/", verify=bundle)
    if r is not None and r.status_code == 200:
        for t, u in links(r.text, r.url):
            if re.search(r"bolet|estad|anuario|\.pdf|\.xls", t + u, re.I):
                out(f"   LINK {t[:100]!r} -> {u}")
    r = get("https://www.mhe.gob.bo/wp-content/uploads/2025/12/Boletin-trimestral-3T_2025-Intranet.pdf", verify=bundle)
    if r is not None and r.status_code == 200:
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if re.search(r"Exportaci[oó]n de gas|Brasil|Argentina|mercado interno de gas|Comercializaci", text, re.I):
                    out(f"  --- page {i + 1}")
                    for ln in text.splitlines()[:60]:
                        out("     ", ln[:230])
except Exception as e:
    out("  MHE failed:", type(e).__name__, str(e)[:300])

out("\n=================== 5. ANH / YPFB Transporte / YPFB gas page")
for page in ("https://www.anh.gob.bo/contenido.php?s=8", "https://www.anh.gob.bo/w2019/contenido.php?s=8",
             "https://www.ypfbtransporte.com.bo/", "https://www.ypfbtransporte.com.bo/estadisticas",
             "https://www.ypfb.gob.bo/Gas_natural"):
    r = get(page)
    if r is None or r.status_code != 200:
        continue
    for t, u in links(r.text, r.url):
        if re.search(r"bolet|estad|export|volum|produc|anuario|\.pdf|\.xls", t + u, re.I):
            out(f"   LINK {t[:100]!r} -> {u}")
    if "Gas_natural" in page:
        txt = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"\s+", " ", txt)
        for m in re.finditer(r"(export|Brasil|Argentina|MMm|millones de metros)", txt, re.I):
            out("   ...", txt[max(0, m.start() - 150): m.start() + 200])
