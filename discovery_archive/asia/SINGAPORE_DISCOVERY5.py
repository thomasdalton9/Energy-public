"""
Singapore discovery, round 5. Round 4 found:
  - NEMS download form: /api/sitecore/DataSync/DataDownload?value=<code>&fromDate=&toDate=&tpcValue=1
    (<=31 days) and /api/sitecore/DataSync/DataDownloadByYear?value=<code>&year=YYYY&tpcValue=1.
    Codes: 1 = USEP and Demand Forecast, 16 = Metered Generation by Facility Type, 5 = Wholesale prices.
  - EMA half-hourly table: /bin/corporate-site/half-hourly-list?csvPath=...&assetFolderPath=...&startYear=2014
This round tests both and lists the other EMA /bin/corporate-site servlets (for the monthly statistics list).
"""
import io
import re
import warnings
import zipfile

import requests

warnings.filterwarnings("ignore")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (10, 180)
NEMS = "https://www.nems.emcsg.com"
EMA = "https://www.ema.gov.sg"


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} "
            f"{r.headers.get('content-disposition')} {len(r.content)} bytes")
        return r
    except requests.RequestException as e:
        out(f"ERR {url}: {e}")
        return None


def show(r, n=12, tail=5):
    if r is None or not r.content:
        return
    c = r.content
    if c[:2] == b"PK":
        try:
            z = zipfile.ZipFile(io.BytesIO(c))
            out(f"  zip: {z.namelist()[:20]}")
            if any(n_.endswith(".csv") for n_ in z.namelist()):
                c = z.read(next(n_ for n_ in z.namelist() if n_.endswith(".csv")))
            else:
                import pandas as pd
                x = pd.read_excel(io.BytesIO(c), sheet_name=None, header=None)
                for s, df in x.items():
                    out(f"  xlsx sheet {s} {df.shape}\n{df.head(n).to_string()[:2500]}\n...\n{df.tail(tail).to_string()[:1200]}")
                return
        except zipfile.BadZipFile:
            pass
    t = c.decode("utf-8-sig", errors="replace")
    lines = t.splitlines()
    out(f"  {len(lines)} lines")
    for l in lines[:n]:
        out(f"   | {l[:300]}")
    out("   ...")
    for l in lines[-tail:]:
        out(f"   | {l[:300]}")


def ema():
    out("\n===== EMA half-hourly servlet =====")
    r = get(f"{EMA}/bin/corporate-site/half-hourly-list", params={
        "csvPath": "/content/dam/corporate/statistics/half-hourly-data.csv",
        "assetFolderPath": "/content/dam/corporate/resources/statistics/half-hourly-data", "startYear": "2014"})
    if r is None:
        return
    try:
        j = r.json()
    except ValueError:
        out(r.text[:2000])
        return
    out(f"  keys {list(j)} years {j.get('years')} months {j.get('months')}")
    items = j.get("items", [])
    out(f"  {len(items)} items")
    for it in items[:5] + items[-3:]:
        out(f"   {it}")
    for it in items[:2]:
        for k in ("Excel", "CSV", "Csv", "PDF"):
            if it.get(k):
                show(get(EMA + it[k] if it[k].startswith("/") else it[k]), 15, 5)
    old = [it for it in items if str(it.get("Date", "")).startswith("2021-01")]
    for it in old[:1]:
        for k in ("Excel", "CSV", "Csv"):
            if it.get(k):
                show(get(EMA + it[k] if it[k].startswith("/") else it[k]), 8, 3)


def ema_servlets():
    out("\n===== EMA /bin servlets in clientlib-site =====")
    r = get(f"{EMA}/resources/statistics")
    js = sorted(set(re.findall(r'src="(/etc\.clientlibs/corporate/clientlibs/clientlib-site[^"]+\.js)"', r.text)))
    t = get(EMA + js[0]).text
    for s in sorted(set(re.findall(r'/bin/[A-Za-z0-9_\-/.]+', t))):
        out(f"  SERVLET {s}")
        for m in list(re.finditer(re.escape(s), t))[:1]:
            out(f"    ctx {t[max(0, m.start() - 600): m.start() + 500]!r}")
    i = r.text.find("statistics-category")
    out(r.text[max(0, i - 5000): i - 2500])


def nems():
    out("\n===== NEMS downloads =====")
    for path, params in [
        ("DataDownload", {"value": "1", "fromDate": "2026-09-01", "toDate": "2026-09-03", "tpcValue": "1"}),
        ("DataDownload", {"value": "16", "fromDate": "2026-09-01", "toDate": "2026-09-03", "tpcValue": "1"}),
        ("DataDownload", {"value": "5", "fromDate": "2026-08-01", "toDate": "2026-08-03", "tpcValue": "1"}),
        ("DataDownloadByYear", {"value": "1", "year": "2021", "tpcValue": "1"}),
        ("DataDownloadByYear", {"value": "16", "year": "2025", "tpcValue": "1"}),
        ("DataDownload", {"value": "1", "fromDate": "2021-01-01", "toDate": "2021-01-02", "tpcValue": "1"}),
        ("DataDownload", {"value": "16", "fromDate": "2021-06-01", "toDate": "2021-06-02", "tpcValue": "1"}),
    ]:
        show(get(f"{NEMS}/api/sitecore/DataSync/{path}", params=params), 12, 4)


if __name__ == "__main__":
    for f in (nems, ema, ema_servlets):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__} failed: {e!r}")
