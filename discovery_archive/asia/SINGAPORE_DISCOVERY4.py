"""
Singapore discovery, round 4. Round 3 found:
  - NEMS (EMC) market data download: GET https://www.nems.emcsg.com/api/sitecore/DataSync/DataDownload
    with value / fromDate / toDate / seriesName / tpcValue / year; CSV, up to 31 days per file, five-year
    rolling window. USEP and Demand Forecast (D+6) etc.
  - EMA half-hourly table component: data-resourcePath=/content/dam/corporate/statistics/half-hourly-data.csv
    (404 directly) and data-assetFolderPath=/content/dam/corporate/resources/statistics/half-hourly-data;
    the Vue code lives in clientlib-site.min.js.
This round: list the NEMS download options and test them; find how the EMA component builds its URLs.
"""
import re
import warnings

import requests

warnings.filterwarnings("ignore")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,text/csv,*/*"}
T = (10, 120)
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


def nems():
    out("\n===== NEMS =====")
    r = get(f"{NEMS}/nems-prices")
    t = r.text
    i = t.find("daily-calendar")
    out(t[max(0, i - 6000): i + 4000])
    for m in re.finditer(r"<select[^>]*>(.*?)</select>", t, re.S):
        opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>(.*?)</option>', m.group(1), re.S)
        out(f"SELECT {m.group(0)[:150]!r}: {opts[:40]}")
    for kw in ["OnSubmitDataDownloadForm", "valuefield", "seriesName", "tabvalue", "DataDownload"]:
        for mm in list(re.finditer(kw, t))[:3]:
            out(f"CTX[{kw}] {t[max(0, mm.start() - 300): mm.start() + 1200]!r}")
    for js in sorted(set(re.findall(r'src="(/[^"]+\.js[^"]*)"', t))):
        jr = get(NEMS + js)
        if jr is None:
            continue
        for kw in ["DataDownload", "seriesName", "valuefield", "tpcValue", "GetChartData", "/api/"]:
            for mm in list(re.finditer(re.escape(kw), jr.text))[:3]:
                out(f"  JS {js[-40:]} [{kw}] {jr.text[max(0, mm.start() - 400): mm.start() + 600]!r}")


def nems_try():
    out("\n===== NEMS download attempts =====")
    for params in [
        {"value": "USEP", "fromDate": "2026-09-01", "toDate": "2026-09-02", "tpcValue": "1"},
        {"value": "Demand", "fromDate": "2026-09-01", "toDate": "2026-09-02", "tpcValue": "1"},
        {"value": "", "fromDate": "2026-09-01", "toDate": "2026-09-02", "tpcValue": "1"},
        {"value": "USEP", "fromDate": "2021-10-05", "toDate": "2021-10-06", "tpcValue": "1"},
        {"value": "USEP", "fromDate": "2026-08-01", "toDate": "2026-08-31", "tpcValue": "1"},
    ]:
        r = get(f"{NEMS}/api/sitecore/DataSync/DataDownload", params=params)
        if r is not None:
            out(r.text[:1500])
            out("   ... " + r.text[-400:])


def ema_js():
    out("\n===== EMA clientlib-site =====")
    r = get(f"{EMA}/resources/statistics/half-hourly-system-demand-data")
    js = sorted(set(re.findall(r'src="(/etc\.clientlibs/corporate/clientlibs/clientlib-site[^"]+\.js)"', r.text)))
    for j in js:
        jr = get(EMA + j)
        t = jr.text
        for kw in ["assetFolderPath", "resourcePath", "corporate-table", "yearsDataHidden", "startYear",
                   "paginatedItems", "item.PDF", "PDF:", "CSV:"]:
            for mm in list(re.finditer(re.escape(kw), t))[:3]:
                out(f"  [{kw}] {t[max(0, mm.start() - 800): mm.start() + 1500]!r}")


if __name__ == "__main__":
    for f in (nems, nems_try, ema_js):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__} failed: {e!r}")
