"""
Bolivia gas demand by sector - round 2.

1. Ministerio de Hidrocarburos y Energias (mhe.gob.bo) serves an
   incomplete certificate chain (no intermediate), so normal clients
   reject it. Fix it properly, with verification ON: read the server's
   leaf certificate, follow its AIA "CA Issuers" URL to download the
   missing intermediate, and verify against certifi's roots plus that
   intermediate. Then dump the gas pages of the quarterly "Boletin
   Energetico" and list the ministry's bulletin links.
2. INE monthly spreadsheets: gas sold through distribution networks
   ("Red de Distribucion") and vehicle gas (GNV) - print their layout.
3. YPFB bulletin list page.
Not reachable from the editing sandbox.
"""
import io
import re
import socket
import ssl
import tempfile
from urllib.parse import urljoin

import certifi
import pandas as pd
import pdfplumber
import requests
from cryptography import x509
from cryptography.hazmat.primitives.serialization import Encoding
from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)


def out(*a):
    print(*a, flush=True)


def bundle_with_intermediates(host):
    """certifi roots + intermediates fetched via the leaf's AIA URLs."""
    pem = ssl.get_server_certificate((host, 443), timeout=30)  # reading the cert only; nothing trusted yet
    leaf = x509.load_pem_x509_certificate(pem.encode())
    out(f"  leaf subject: {leaf.subject.rfc4514_string()} | issuer: {leaf.issuer.rfc4514_string()}")
    extra = []
    cert = leaf
    for _ in range(3):
        try:
            aia = cert.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
        except x509.ExtensionNotFound:
            break
        urls = [d.access_location.value for d in aia if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS]
        if not urls:
            break
        out(f"  AIA CA issuers: {urls}")
        raw = requests.get(urls[0], headers=H, timeout=T).content
        try:
            cert = x509.load_der_x509_certificate(raw)
        except ValueError:
            cert = x509.load_pem_x509_certificate(raw)
        out(f"  fetched intermediate: {cert.subject.rfc4514_string()}")
        extra.append(cert.public_bytes(Encoding.PEM).decode())
        if cert.subject == cert.issuer:
            break
    f = tempfile.NamedTemporaryFile("w", suffix=".pem", delete=False)
    f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    f.close()
    return f.name


def pdf_dump(content, page_pat, max_lines=80, max_pages=10):
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        shown = 0
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not re.search(page_pat, text, re.I):
                continue
            shown += 1
            if shown > max_pages:
                break
            out(f"  --- page {i + 1}")
            for ln in text.splitlines()[:max_lines]:
                out("     ", ln[:200])
            for tb in page.extract_tables()[:2]:
                out(f"     TABLE {len(tb)} rows")
                for row in tb[:25]:
                    out("       ", [str(x)[:14].replace("\n", " ") if x else "" for x in row][:14])


out("=================== 1. MHE (certificate chain fix)")
try:
    bundle = bundle_with_intermediates("www.mhe.gob.bo")
    for u in ("https://www.mhe.gob.bo/wp-content/uploads/2025/12/Boletin-trimestral-3T_2025-Intranet.pdf",):
        r = requests.get(u, headers=H, timeout=T, verify=bundle)
        out(f"GET {u} -> {r.status_code} {len(r.content)}B")
        if r.status_code == 200 and r.content[:4] == b"%PDF":
            pdf_dump(r.content, r"gas natural|mercado interno|termoel|GNV|consumo", max_pages=12)
    for page in ("https://www.mhe.gob.bo/", "https://www.mhe.gob.bo/boletines/", "https://www.mhe.gob.bo/?s=boletin"):
        r = requests.get(page, headers=H, timeout=T, verify=bundle)
        out(f"GET {page} -> {r.status_code}")
        if r.status_code == 200:
            for h in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
                if re.search(r"bolet|estad|anuario|\.pdf", h, re.I):
                    out("   ", h[:170])
except Exception as e:
    out("  MHE failed:", type(e).__name__, str(e)[:300])

out("\n=================== 2. INE spreadsheets")
for label, u in (("Red de distribucion", "https://nube.ine.gob.bo/index.php/s/Jsm9uMT7L9ALmWr/download"),
                 ("GNV por departamento", "https://nube.ine.gob.bo/index.php/s/oAibslam6h0MHq2/download"),
                 ("Produccion gas por departamento", "https://nube.ine.gob.bo/index.php/s/QQ1WOj4zNwoeFlx/download")):
    try:
        r = requests.get(u, headers=H, timeout=T)
        out(f"\n  ## {label}: {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        for name, df in list(pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items())[:2]:
            out(f"   sheet {name!r} {df.shape}")
            for row in df.head(18).itertuples(index=False):
                out("      ", [str(x)[:16] for x in row if str(x) != "nan"][:16])
            out("      ...")
            for row in df.tail(4).itertuples(index=False):
                out("      ", [str(x)[:16] for x in row if str(x) != "nan"][:16])
    except Exception as e:
        out(f"  {label}: ERR {type(e).__name__}: {str(e)[:150]}")

out("\n=================== 3. YPFB bulletins")
try:
    r = requests.get("https://www.ypfb.gob.bo/pagina-lista-boletines", headers=H, timeout=T)
    out(f"  status {r.status_code}")
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(r"bolet|estad|\.pdf|mercado|gas", h + t, re.I):
            out(f"   {t[:80]!r} -> {urljoin('https://www.ypfb.gob.bo/', h)}")
except Exception as e:
    out("  YPFB failed:", e)
