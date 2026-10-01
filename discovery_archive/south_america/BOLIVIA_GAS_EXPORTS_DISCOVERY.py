"""
Bolivia gas production and exports (Brazil / GASBOL, Argentina / GJA),
monthly from 2021 - source search.

1. INE hydrocarbons page: list every table link, then dump the layout of
   each gas table (production by department H02, exports, etc.).
2. INE external-trade pages: exports of natural gas by country of
   destination (volume / value).
3. YPFB: 'Gas natural' page, bulletin listings, the latest 'Boletin
   informativo' PDFs (production / export pages).
4. ANH: statistics pages.
5. MHE quarterly 'Boletin Energetico' (incomplete certificate chain:
   fetch the missing intermediate via AIA and verify against certifi).
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
        res.append((t, urljoin(base, h.replace("&amp;", "&"))))
    return res


def dump_xlsx(content, head=14, tail=8):
    try:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:
        out(f"   not a spreadsheet: {type(e).__name__} {str(e)[:100]}")
        return
    for name, df in list(sheets.items())[:3]:
        out(f"   sheet {name!r} {df.shape}")
        for row in df.head(head).itertuples(index=False):
            out("      ", [str(x)[:60] for x in row if str(x) != "nan"][:14])
        out("       ...")
        for row in df.tail(tail).itertuples(index=False):
            out("      ", [str(x)[:60] for x in row if str(x) != "nan"][:14])


def pdf_dump(content, page_pat, max_lines=70, max_pages=8):
    try:
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
                    out("     ", ln[:220])
    except Exception as e:
        out("  pdf failed:", type(e).__name__, str(e)[:150])


GAS = re.compile(r"gas", re.I)

out("=================== 1. INE hydrocarbons tables")
ine_pages = ["https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/"
             "hidrocarburo-cuadros-estadisticos/",
             "https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/"
             "hidrocarburos-introduccion/"]
seen = set()
for page in ine_pages:
    r = get(page)
    if not r or r.status_code != 200:
        continue
    for t, u in links(r.text, page):
        if "nube.ine" in u or re.search(r"\.xlsx?$|download", u):
            out(f"   LINK {t[:110]!r} -> {u}")
            if GAS.search(t) and u not in seen:
                seen.add(u)
                rr = get(u)
                if rr is not None and rr.status_code == 200:
                    dump_xlsx(rr.content)

out("\n=================== 2. INE external trade (exports by product / country)")
for page in ("https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/"
             "cuadros-estadisticos-exportaciones/",
             "https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/",
             "https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/exportaciones/"):
    r = get(page)
    if not r or r.status_code != 200:
        continue
    for t, u in links(r.text, page):
        if "nube.ine" in u or re.search(r"\.xlsx?$|download|export", u + t, re.I):
            out(f"   LINK {t[:110]!r} -> {u}")
            if re.search(r"gas|hidrocarb|producto|pa[ií]s", t, re.I) and u not in seen and "nube.ine" in u:
                seen.add(u)
                rr = get(u)
                if rr is not None and rr.status_code == 200:
                    dump_xlsx(rr.content, head=16, tail=4)

out("\n=================== 3. YPFB")
for page in ("https://www.ypfb.gob.bo/Gas_natural", "https://www.ypfb.gob.bo/pagina-lista-boletines",
             "https://www.ypfb.gob.bo/", "https://www.ypfb.gob.bo/boletin-estadistico",
             "https://www.ypfb.gob.bo/Boletin_estadistico", "https://www.ypfb.gob.bo/estadisticas"):
    r = get(page)
    if not r or r.status_code != 200:
        continue
    for t, u in links(r.text, page):
        if re.search(r"bolet|estad|export|produc|\.pdf|\.xls", t + u, re.I):
            out(f"   LINK {t[:90]!r} -> {u}")
for u in ("https://www.ypfb.gob.bo/sites/default/files/2026-09/BOLETIN%20INFORMATIVO_8_2026.pdf",
          "https://www.ypfb.gob.bo/sites/default/files/2026-09/BOLETIN%20INFORMATIVO_4.pdf"):
    r = get(u)
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        pdf_dump(r.content, r"export|producci[oó]n|MMm3|MMmcd|Brasil|Argentina", max_pages=5)

out("\n=================== 4. ANH")
for page in ("https://www.anh.gob.bo/", "https://www.anh.gob.bo/w2019/contenido.php?s=13",
             "https://www.anh.gob.bo/w2019/estadisticas.php", "https://www.anh.gob.bo/InsideFiles/Inicio/"):
    r = get(page)
    if not r or r.status_code != 200:
        continue
    for t, u in links(r.text, page):
        if re.search(r"bolet|estad|export|produc|\.pdf|\.xls|gas", t + u, re.I):
            out(f"   LINK {t[:90]!r} -> {u}")


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
        out(f"  intermediate: {cert.subject.rfc4514_string()}")
        if cert.subject == cert.issuer:
            break
    f = tempfile.NamedTemporaryFile("w", suffix=".pem", delete=False)
    f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    f.close()
    return f.name


out("\n=================== 5. MHE Boletin Energetico")
try:
    bundle = bundle_with_intermediates("www.mhe.gob.bo")
    r = get("https://www.mhe.gob.bo/wp-content/uploads/2025/12/Boletin-trimestral-3T_2025-Intranet.pdf",
            verify=bundle)
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        pdf_dump(r.content, r"export|producci[oó]n de gas|Brasil|Argentina", max_pages=8)
    for page in ("https://www.mhe.gob.bo/boletines/", "https://www.mhe.gob.bo/"):
        r = get(page, verify=bundle)
        if r is not None and r.status_code == 200:
            for t, u in links(r.text, page):
                if re.search(r"bolet|estad|anuario|\.pdf", t + u, re.I):
                    out(f"   LINK {t[:90]!r} -> {u}")
except Exception as e:
    out("  MHE failed:", type(e).__name__, str(e)[:300])
