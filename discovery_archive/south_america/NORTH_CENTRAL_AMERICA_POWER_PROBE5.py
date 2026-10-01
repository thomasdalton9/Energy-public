"""
Round 5 (Oct-2026) after NORTH_CENTRAL_AMERICA_POWER_PROBE4.py found:
  AMM  GraficaPW /graficaCombustible?dt=dd/mm/yyyy returns hourly MW by fuel
       (AGUA, BIOGAS, BIOMASA, BUNKER, CARBON, CARBON/PETCOKE, DIESEL, GAS
       NATURAL, IRRADIACION, ...) for 2025-06 and 2026, but [] for 2023-03
       and 2021-01. Find the first date it has, and compare one day with the
       daily Posdespacho 'Carga Horaria' sheet (hourly MW per unit).
  UT   the www IP (190.120.15.116) times out on 80/443 from the runner, like
       its name servers: geo-fenced to El Salvador.
  ODS  the APEX 'Listado website' report rows are not in the page HTML
       (lazy-loaded region); the 'otr' page has today's generation by
       technology only.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE5.py amm|ods|ut
"""

print("STARTING", flush=True)

import datetime as dt
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
        print(f"[{r.status_code}] {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:300]}", flush=True)
        return None


def cells(df, rows=range(0, 60), max_len=30):
    for i in rows:
        if i >= len(df):
            break
        vals = [(j, str(v)[:max_len]) for j, v in enumerate(df.iloc[i].tolist()) if str(v) not in ("nan", "NaT", "None", "")]
        if vals:
            print(f"   r{i}:", " | ".join(f"c{j}={v}" for j, v in vals), flush=True)


# ---------------- AMM ----------------
AMM = "https://wl12.amm.org.gt/GraficaPW/graficaCombustible"


def amm_day(d):
    r = S.get(AMM, params={"dt": d.strftime("%d/%m/%Y")}, timeout=60)
    try:
        return r.json()
    except ValueError:
        return None


def amm():
    import pandas as pd
    lo, hi = dt.date(2023, 3, 10), dt.date(2025, 6, 15)
    while (hi - lo).days > 1:
        mid = lo + (hi - lo) // 2
        data = amm_day(mid)
        print(f"  {mid}: {len(data) if data else 0} rows", flush=True)
        if data:
            hi = mid
        else:
            lo = mid
    print("  FIRST graficaCombustible day:", hi, flush=True)
    for d in [hi + dt.timedelta(days=k) for k in (1, 30, 200)]:
        print("   check", d, len(amm_day(d) or []), flush=True)
    d = dt.date(2026, 9, 25)
    data = amm_day(d)
    df = pd.DataFrame(data)
    df["potencia"] = pd.to_numeric(df["potencia"], errors="coerce")
    print("  hours:", sorted(df["hora"].unique(), key=lambda x: float(x))[:30], flush=True)
    print("  daily MWh by tipo (sum of hourly MW):", flush=True)
    print(df.groupby("tipo")["potencia"].agg(["sum", "count"]).round(1).to_string(), flush=True)
    r = get("https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/2026/09_SEPTIEMBRE/PD20260925.zip")
    z = zipfile.ZipFile(io.BytesIO(r.content))
    sheets = pd.read_excel(io.BytesIO(z.read(z.namelist()[0])), sheet_name=None, header=None)
    ch = sheets["Carga Horaria"]
    units = [str(v) for v in ch.iloc[5].tolist()]
    tot = pd.to_numeric(ch.iloc[6], errors="coerce")
    by_suffix = {}
    for u, v in zip(units, tot):
        if u in ("nan",) or pd.isna(v):
            continue
        s = re.sub(r"\d+$", "", u.split("-")[-1]) if "-" in u else u
        by_suffix[s] = by_suffix.get(s, 0) + v
    print("  posdespacho 25-Sep row6 by unit suffix:", {k: round(v, 1) for k, v in by_suffix.items()}, flush=True)
    print("  Carga Horaria legend rows:", flush=True)
    cells(ch, range(48, 52), 200)
    print("  == Curva S.N.I.", flush=True)
    cells(sheets["Curva S.N.I."], range(0, 46), 25)
    r = get("https://www.amm.org.gt/pdfs2/post_despacho/POSDESPACHO_DIARIO/2021/01_ENERO/PD20210101.zip")
    z = zipfile.ZipFile(io.BytesIO(r.content))
    s21 = pd.read_excel(io.BytesIO(z.read(z.namelist()[0])), sheet_name=None, header=None)
    print("  2021 sheets:", list(s21), flush=True)
    for k in s21:
        if "carga" in k.lower():
            c = s21[k]
            for i in range(0, 8):
                print(f"   2021 {k} r{i}:", [str(v)[:8] for v in c.iloc[i].tolist()][:200], flush=True)


# ---------------- ODS ----------------
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


def ods():
    S.verify = aia_bundle("cnd.enee.hn")
    base = "https://appcnd.enee.hn:3200/odsprd/"
    r = get(base + "f?p=110:6:::::p6_id:2")
    t = r.text
    for m in re.finditer(r"<script[^>]*>(.*?)</script>", t, re.S):
        s = m.group(1).strip()
        if s:
            print("  SCRIPT:", re.sub(r"\s+", " ", s)[:2500], flush=True)
    for m in re.finditer(r"<input[^>]*type=\"hidden\"[^>]*>", t):
        print("  HIDDEN", m.group(0)[:200], flush=True)
    body = t[t.find("<body"):]
    body = re.sub(r"<script.*?</script>", "", body, flags=re.S)
    print("  BODY (tags kept, 6000 chars):", re.sub(r"\s+", " ", body)[:6000], flush=True)
    r = get(base + "ods_prd/r/operador-del-sistema-ods/otr")
    if r is not None:
        for m in re.finditer(r"<script[^>]*>(.*?)</script>", r.text, re.S):
            s = m.group(1).strip()
            if "ajax" in s.lower() or "region" in s.lower():
                print("  OTR SCRIPT:", re.sub(r"\s+", " ", s)[:3000], flush=True)
    get("https://qacnd.enee.hn:3200/odsqa/ods_qa/r/scada-reports153172/home")


def ut():
    for host, port in [("190.120.15.116", 443), ("190.120.15.116", 80), ("190.120.15.100", 53),
                       ("190.242.123.216", 53), ("190.242.150.196", 53)]:
        s = socket.socket()
        s.settimeout(15)
        try:
            s.connect((host, port))
            print(f"  TCP {host}:{port} open", flush=True)
        except OSError as e:
            print(f"  TCP {host}:{port} failed: {e}", flush=True)
        finally:
            s.close()


if __name__ == "__main__":
    {"amm": amm, "ods": ods, "ut": ut}[sys.argv[1]]()
    print("DONE", flush=True)
