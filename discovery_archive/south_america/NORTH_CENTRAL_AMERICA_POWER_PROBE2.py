"""
Round 2 of NORTH_CENTRAL_AMERICA_POWER_DISCOVERY.py (Oct-2026).

Round 1 found:
  AMM  wl12.amm.org.gt/GraficaPW/tiempoRealDia.jsp ("Grafica Generacion AMM",
       by technology / resource, Excel export) backed by graficaCombustible
       (JSON; ?dt=dd/mm/yyyy returned []), and daily Posdespacho zips
       pdfs2/post_despacho/POSDESPACHO_DIARIO/<yyyy>/<mm_MES>/PD<yyyymmdd>.zip
  UT   www.ut.com.sv / estadistico.ut.com.sv did not resolve from the runner
  ODS  www.ods.org.hn redirects to cnd.enee.hn, whose server omits its
       intermediate certificate (verify fails "unable to get local issuer")
  EOR  informes-diarios-de-operacion, dashboards, mapa/map.html, inf_est.pdf
  BEL  only an annual report PDF

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE2.py amm|ut|ods|eor
"""

print("STARTING", flush=True)

import io
import json
import re
import socket
import ssl
import sys
import zipfile
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9"}
S = requests.Session()
S.headers.update(H)


def get(url, **kw):
    try:
        r = S.get(url, timeout=60, **kw)
        print(f"[{r.status_code}] {r.url} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:300]}", flush=True)
        return None


def show_excel(content, name, rows=12, width=14):
    import pandas as pd
    try:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:  # noqa: BLE001
        print(f"  cannot read {name}: {e}", flush=True)
        return
    for sn, df in sheets.items():
        print(f"  -- {name} / sheet {sn!r} shape={df.shape}", flush=True)
        for i in range(min(rows, len(df))):
            vals = [str(v)[:22] for v in df.iloc[i, :width].tolist()]
            print("     ", " | ".join(vals), flush=True)


def links(html, base, pat=None):
    out = []
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", html, re.I | re.S):
        href = urljoin(base, m.group(1).strip())
        text = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:80]
        if pat is None or re.search(pat, href + " " + text, re.I):
            out.append((href, text))
    return out


# ---------------- AMM ----------------
def amm():
    base = "https://wl12.amm.org.gt/GraficaPW/"
    for js in ["js/tiempoRealDia.js", "js/loaderChart.js"]:
        r = get(base + js)
        if r is not None and r.ok:
            t = r.text
            print(f"---- {js} ({len(t)} chars)", flush=True)
            print(t[:9000], flush=True)
    r = get(base + "tiempoRealDia.jsp")
    if r is not None:
        print(r.text[:4000], flush=True)
    # daily posdespacho zip
    for url in ["https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/2026/09_SEPTIEMBRE/PD20260925.zip",
                "https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/2021/01_ENERO/PD20210101.zip"]:
        r = get(url)
        if r is None or not r.ok or not r.content[:2] == b"PK":
            continue
        z = zipfile.ZipFile(io.BytesIO(r.content))
        for info in z.infolist():
            print(f"  ZIP {info.filename} {info.file_size:,}", flush=True)
        if "2026" in url:
            for info in z.infolist():
                if re.search(r"\.xls[xm]?$", info.filename, re.I):
                    show_excel(z.read(info), info.filename, rows=10, width=12)
    r = get("https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_MENSUAL/2026/IPM202608.zip")
    if r is not None and r.content[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(r.content))
        for info in z.infolist():
            print(f"  ZIP {info.filename} {info.file_size:,}", flush=True)
    r = get("https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/pubamm/2026/GM20260917.xlsx")
    if r is not None and r.ok:
        show_excel(r.content, "GM20260917.xlsx", rows=15, width=16)
    r = get("https://wl12.amm.org.gt/reportesDinamicosMM/")
    if r is not None:
        for h, t in links(r.text, r.url)[:80]:
            print("  LINK", h, "|", t, flush=True)
        for m in re.finditer(r"<(form|select|input|iframe)[^>]*>", r.text, re.I):
            print("  ", m.group(0)[:200], flush=True)


# ---------------- UT ----------------
def ut():
    for host in ["ut.com.sv", "www.ut.com.sv", "estadistico.ut.com.sv", "siget.gob.sv", "www.siget.gob.sv"]:
        try:
            print(host, "->", socket.gethostbyname(host), flush=True)
        except OSError as e:
            print(host, "-> resolve failed:", e, flush=True)
        try:
            d = requests.get("https://dns.google/resolve", params={"name": host, "type": "A"}, timeout=20).json()
            print("   DoH:", d.get("Status"), [a.get("data") for a in d.get("Answer", [])],
                  [a.get("data") for a in d.get("Authority", [])][:2], flush=True)
        except Exception as e:  # noqa: BLE001
            print("   DoH failed", e, flush=True)
    for url in ["https://ut.com.sv/", "http://www.ut.com.sv/", "http://estadistico.ut.com.sv/OperacionDiaria.aspx"]:
        get(url)


# ---------------- ODS (missing intermediate) ----------------
def aia_bundle(host):
    """certifi roots + the intermediate(s) named in the server cert's AIA 'CA Issuers' URL.
    Verification stays fully on; this only supplies the chain piece the server fails to send."""
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, 443))
    cert = x509.load_pem_x509_certificate(pem.encode())
    print("  leaf subject:", cert.subject.rfc4514_string(), "\n  issuer:", cert.issuer.rfc4514_string(),
          "\n  valid:", cert.not_valid_before_utc, cert.not_valid_after_utc, flush=True)
    extra = []
    for _ in range(3):
        try:
            aia = cert.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
        except x509.ExtensionNotFound:
            break
        urls = [d.access_location.value for d in aia if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS]
        if not urls:
            break
        print("  AIA:", urls, flush=True)
        der = requests.get(urls[0], timeout=30).content
        try:
            cert = x509.load_der_x509_certificate(der)
        except ValueError:
            cert = x509.load_pem_x509_certificate(der)
        print("  intermediate:", cert.subject.rfc4514_string(), flush=True)
        extra.append(cert.public_bytes(Encoding.PEM).decode())
        if cert.issuer == cert.subject:
            break
    path = f"/tmp/{host}_bundle.pem"
    open(path, "w").write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    return path


def ods():
    bundle = aia_bundle("cnd.enee.hn")
    S.verify = bundle
    seen, queue, n = set(), ["https://cnd.enee.hn/", "https://www.ods.org.hn/"], 0
    pat = r"informe|report|estad|operac|generac|produc|datos|despacho|diari|mensual|anual|tiempo|hist|xls|csv|json|api|pdf"
    data = {}
    while queue and n < 30:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        r = get(u)
        n += 1
        if r is None or "html" not in (r.headers.get("content-type") or ""):
            continue
        t = r.text
        title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        print("   title:", title.group(1).strip()[:100] if title else "", flush=True)
        for m in re.finditer(r"<(iframe|form)[^>]*>", t, re.I):
            print("   ", m.group(0)[:200], flush=True)
        for m in re.finditer(r"""["'`]((?:https?:)?/[^"'`\s]*(?:api|json|ajax|\.php\?|ashx|svc|data)[^"'`\s]*)["'`]""", t, re.I):
            print("    APIREF", m.group(1)[:200], flush=True)
        for h, tx in links(t, r.url, pat):
            if re.search(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)", h, re.I):
                data[h] = tx
            elif "enee.hn" in h or "ods.org.hn" in h:
                queue.append(h)
    for h, tx in list(data.items())[:120]:
        print("  DATA", h, "|", tx, flush=True)
    print("  QUEUE", queue[:40], flush=True)


# ---------------- EOR ----------------
def eor():
    for u in ["https://enteoperador.org/mer/gestion-tecnica-operativa/informes-diarios-de-operacion/",
              "https://enteoperador.org/dashboards/",
              "https://enteoperador.org/mer/gestion-comercial/informes-de-transaccion-diario-mensual-y-anual/",
              "https://enteoperador.org/mer/publicaciones-historicas/"]:
        r = get(u)
        if r is None:
            continue
        t = r.text
        main = t[t.find("<main") if "<main" in t else 0:]
        for m in re.finditer(r"<(iframe|embed|object)[^>]*>", main, re.I):
            print("   ", m.group(0)[:250], flush=True)
        for h, tx in links(main, r.url):
            if re.search(r"facebook|twitter|linkedin|youtube|instagram|/category/|/tag/|#", h):
                continue
            if re.search(r"diari|operac|posdesp|estad|dashboard|generac|xls|csv|pdf|zip|json|informe|power|app\.", h + tx, re.I):
                print("   LINK", h, "|", tx, flush=True)
    r = get("https://enteoperador.org/mapa/map.html")
    if r is not None:
        print(r.text[:3500], flush=True)
        for m in re.finditer(r"<script[^>]*src=[\"']([^\"']+)", r.text, re.I):
            js = urljoin(r.url, m.group(1))
            if "jquery" in js or "leaflet" in js.lower():
                continue
            j = get(js)
            if j is not None and j.ok:
                for mm in re.finditer(r"""["'`]([^"'`\s]*(?:api|json|\.php|ashx|svc|http)[^"'`\s]*)["'`]""", j.text, re.I):
                    print("    JSREF", mm.group(1)[:200], flush=True)
    r = get("https://enteoperador.org/archivos/document/inf_est_diario_RMER/Informes_Estadisticos/inf_est.pdf")
    if r is not None and r.content[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print(f"  inf_est.pdf pages={len(pdf.pages)}", flush=True)
            for i, p in enumerate(pdf.pages[:6]):
                print(f"  --- page {i + 1}", flush=True)
                print((p.extract_text() or "")[:2500], flush=True)


if __name__ == "__main__":
    {"amm": amm, "ut": ut, "ods": ods, "eor": eor}[sys.argv[1]]()
    print("DONE", flush=True)
