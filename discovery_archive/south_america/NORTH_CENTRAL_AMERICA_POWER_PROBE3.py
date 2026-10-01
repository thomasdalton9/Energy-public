"""
Round 3 (Oct-2026) after NORTH_CENTRAL_AMERICA_POWER_PROBE2.py found:
  AMM  GraficaPW page loads js/graph.js (the web-service calls);
       daily Posdespacho xlsx has 'Carga Horaria' (hourly MW per unit) and
       'RSO1..3' operation summary pages; GM<date>.xlsx = monthly GWh per
       plant grouped by technology.
  UT   ut.com.sv returns SERVFAIL from Google DNS too.
  ODS  cnd.enee.hn intermediate = Sectigo DV R36 (AIA), root R46 via .p7c.
  EOR  only MER (cross-border) transactions; SCADA map web service.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE3.py amm|ut|ods|eor
"""

print("STARTING", flush=True)

import io
import re
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


def cells(df, max_rows=80, max_len=40):
    """Print non-empty cells row by row."""
    for i in range(min(max_rows, len(df))):
        vals = [(j, str(v)[:max_len]) for j, v in enumerate(df.iloc[i].tolist()) if str(v) not in ("nan", "NaT", "None", "")]
        if vals:
            print(f"   r{i}:", " | ".join(f"c{j}={v}" for j, v in vals), flush=True)


# ---------------- AMM ----------------
def amm():
    import pandas as pd
    r = get("https://wl12.amm.org.gt/GraficaPW/js/graph.js")
    eps = set()
    if r is not None and r.ok:
        t = r.text
        print(f"---- graph.js {len(t)} chars", flush=True)
        for m in re.finditer(r".{0,160}(uri|WebService|xhr|ajax|url|http|fetch|\.get\(|Rest|webresources).{0,200}", t):
            print("  JS:", m.group(0).strip()[:380], flush=True)
        for m in re.finditer(r"""["']([A-Za-z]+(?:/[A-Za-z]+)*/)["']\s*\+\s*fecha""", t):
            eps.add(m.group(1))
        m = re.search(r"var\s+uri\s*=\s*([^;]+);", t)
        print("  uri =", m.group(1) if m else None, flush=True)
    # posdespacho: unit columns of 'Carga Horaria' and the RSO summary pages
    r = get("https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/2026/09_SEPTIEMBRE/PD20260925.zip")
    if r is not None and r.content[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(r.content))
        x = z.read(z.namelist()[0])
        sheets = pd.read_excel(io.BytesIO(x), sheet_name=None, header=None)
        ch = sheets["Carga Horaria"]
        print("  Carga Horaria units:", [str(v) for v in ch.iloc[5].tolist()], flush=True)
        print("  Carga Horaria row6 :", [str(v)[:8] for v in ch.iloc[6].tolist()], flush=True)
        print("  Carga Horaria col0 :", [str(v)[:10] for v in ch.iloc[:, 0].tolist()], flush=True)
        for i in range(30, len(ch)):
            print("   tail", i, [str(v)[:10] for v in ch.iloc[i, :12].tolist()], flush=True)
        for sn in ["RSO1", "RSO2"]:
            print(f"  == {sn}", flush=True)
            cells(sheets[sn], 70)
    # GM: full first column (plant names + group headers)
    r = get("https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/pubamm/2026/GM20260917.xlsx")
    if r is not None and r.ok:
        df = pd.read_excel(io.BytesIO(r.content), sheet_name="Energia", header=None)
        for i in range(len(df)):
            v = str(df.iloc[i, 1])
            if v != "nan":
                print(f"   GM r{i}: {v[:45]} | jan={str(df.iloc[i, 2])[:8]} aug={str(df.iloc[i, 9])[:8]}", flush=True)
    for y in ["2021", "2022"]:
        get(f"https://www.amm.org.gt/pdfs2/pub_gen_mensual_x_planta/pubamm/{y}/GM{y}1231.xlsx")
    # GraficaPW web service guesses
    base = "https://wl12.amm.org.gt/GraficaPW/"
    for ep in sorted(eps) + ["graficaCombustible", "graficaTecnologia", "graficaRecurso"]:
        for d in ["25/09/2026", "25-09-2026", "2026-09-25"]:
            for url in [base + ep.strip("/") + "/" + d, base + ep.strip("/") + "?fecha=" + d]:
                r = get(url)
                if r is not None and r.ok and len(r.content) > 5:
                    print("   ", r.text[:600], flush=True)


# ---------------- UT ----------------
def ut():
    for name, typ in [("ut.com.sv", "A"), ("ut.com.sv", "NS"), ("www.ut.com.sv", "A"), ("com.sv", "NS"),
                      ("ut.sv", "A"), ("www.ut.sv", "A"), ("cne.gob.sv", "A"), ("www.cne.gob.sv", "A")]:
        for prov, url in [("google-cd", "https://dns.google/resolve"), ("cloudflare", "https://cloudflare-dns.com/dns-query")]:
            try:
                p = {"name": name, "type": typ}
                if prov == "google-cd":
                    p["cd"] = "1"
                d = requests.get(url, params=p, headers={"accept": "application/dns-json"}, timeout=20).json()
                print(f"{name} {typ} {prov}: status={d.get('Status')} ans={[a.get('data') for a in d.get('Answer', [])][:4]} "
                      f"auth={[a.get('data') for a in d.get('Authority', [])][:2]} comment={d.get('Comment')}", flush=True)
            except Exception as e:  # noqa: BLE001
                print(name, typ, prov, "failed", e, flush=True)
    for u in ["https://www.cne.gob.sv/", "https://cne.gob.sv/", "https://www.siget.gob.sv/"]:
        r = get(u)
        if r is not None and r.ok:
            for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", r.text, re.I | re.S):
                h, tx = m.group(1), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
                if re.search(r"estad|generac|boletin|electric|mercado|datos|transacc|ut\.com", h + tx, re.I):
                    print("   LINK", urljoin(r.url, h), "|", tx[:70], flush=True)


# ---------------- ODS ----------------
def aia_bundle(host):
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, 443))  # read the leaf only to find its AIA URL
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
        if urls[0].endswith(".p7c"):
            certs = pkcs7.load_der_pkcs7_certificates(der)
        else:
            certs = [x509.load_der_x509_certificate(der)]
        for c in certs:
            print("  chain cert:", c.subject.rfc4514_string(), "<-", c.issuer.rfc4514_string(), flush=True)
            extra.append(c.public_bytes(Encoding.PEM).decode())
        cert = certs[0]
        if cert.issuer == cert.subject:
            break
    path = f"/tmp/{host}_bundle.pem"
    open(path, "w").write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    return path


def ods():
    S.verify = aia_bundle("cnd.enee.hn")
    seen, queue, n, data = set(), ["https://cnd.enee.hn/", "https://www.ods.org.hn/"], 0, {}
    pat = r"informe|report|estad|operac|generac|produc|datos|despacho|diari|mensual|anual|tiempo|hist|xls|csv|json|api|pdf|energ"
    while queue and n < 35:
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
        for m in re.finditer(r"""["'`]((?:https?:)?/[^"'`\s<>]*(?:api|json|ajax|\.php\?|ashx|svc|powerbi)[^"'`\s<>]*)["'`]""", t, re.I):
            print("    APIREF", m.group(1)[:200], flush=True)
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", t, re.I | re.S):
            h = urljoin(r.url, m.group(1).strip())
            tx = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:80]
            if not re.search(pat, h + " " + tx, re.I):
                continue
            if re.search(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)", h, re.I):
                data[h] = tx
            elif ("enee.hn" in h or "ods.org.hn" in h) and h not in seen:
                queue.append(h)
    for h, tx in list(data.items())[:150]:
        print("  DATA", h, "|", tx, flush=True)
    print("  QUEUE", queue[:40], flush=True)


# ---------------- EOR ----------------
def eor():
    r = get("https://enteoperador.org/mapa/js/inicio.js")
    if r is not None:
        for m in re.finditer(r".{0,200}(mapa\.enteoperador\.org|webresources).{0,300}", r.text):
            print("  JS:", m.group(0)[:500], flush=True)
    for u in ["https://mapa.enteoperador.org/WebServiceScadaEORRest/webresources/generic",
              "https://mapa.enteoperador.org/WebServiceScadaEORRest/webresources/generic/"]:
        r = get(u)
        if r is not None:
            print("   ", r.text[:2500], flush=True)


if __name__ == "__main__":
    {"amm": amm, "ut": ut, "ods": ods, "eor": eor}[sys.argv[1]]()
    print("DONE", flush=True)
