"""
Singapore discovery, round 6. Round 5 found:
  - NEMS CSVs work: value=1 (USEP + DEMAND MW, half-hourly), value=16 (metered generation by facility type,
    MWh per half hour); DataDownloadByYear redirects to
    https://emcprodmarketdata.blob.core.windows.net/yearly-market-data-downloads/<USEP|MG>_from_01-Jan-YYYY_to_31-Dec-YYYY.zip
    (one CSV per month inside).
  - EMA half-hourly system demand: /bin/corporate-site/half-hourly-list -> 651 weekly .xls files 2014..2026.
  - EMA statistics page uses /bin/corporate-site/statistics-list?searchString=&category=&pagePath=...
This round: parse the EMA weekly xls layout, test the current-year NEMS yearly zips, and list EMA statistics.
"""
import io
import re
import warnings

import pandas as pd
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
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes "
            f"history={[h.status_code for h in r.history]}")
        return r
    except requests.RequestException as e:
        out(f"ERR {url}: {e}")
        return None


def ema_xls():
    out("\n===== EMA weekly xls layout =====")
    r = get(f"{EMA}/bin/corporate-site/half-hourly-list", params={
        "csvPath": "/content/dam/corporate/statistics/half-hourly-data.csv",
        "assetFolderPath": "/content/dam/corporate/resources/statistics/half-hourly-data", "startYear": "2014"})
    items = r.json()["items"]
    pick = [items[0]] + [it for it in items if it["Date"].endswith("2021")][:1] + \
           [it for it in items if it["Date"].endswith("2021")][-1:] + [items[-1]]
    dates = [it["Date"] for it in items]
    out(f"dates sample {dates[:3]} ... {dates[-3:]}; per year: "
        f"{pd.Series([d[-4:] for d in dates]).value_counts().sort_index().to_dict()}")
    for it in pick:
        x = get(EMA + it["Excel"])
        if x is None or x.status_code != 200:
            continue
        sheets = pd.read_excel(io.BytesIO(x.content), sheet_name=None, header=None)
        for s, df in sheets.items():
            out(f"  {it['Date']} sheet {s!r} shape {df.shape}")
            out(df.head(14).to_string(max_colwidth=22)[:3500])
            out(df.tail(4).to_string(max_colwidth=22)[:1500])


def nems_year():
    out("\n===== NEMS yearly zips =====")
    for v, y in [("1", "2026"), ("16", "2026"), ("16", "2021"), ("1", "2025")]:
        r = get(f"{NEMS}/api/sitecore/DataSync/DataDownloadByYear", params={"value": v, "year": y, "tpcValue": "1"})
        if r is not None and r.content[:2] == b"PK":
            import zipfile
            z = zipfile.ZipFile(io.BytesIO(r.content))
            out(f"  {sorted(z.namelist())}")
            last = sorted(z.namelist(), key=lambda n: pd.to_datetime(n.split("_")[-1].replace(".csv", ""),
                                                                       format="%b-%Y"))[-1]
            t = z.read(last).decode("utf-8-sig", errors="replace").splitlines()
            out(f"  {last}: {len(t)} lines; first {t[:2]}; last {t[-1:]}")
            first = sorted(z.namelist(), key=lambda n: pd.to_datetime(n.split("_")[-1].replace(".csv", ""),
                                                                        format="%b-%Y"))[0]
            t = z.read(first).decode("utf-8-sig", errors="replace").splitlines()
            out(f"  {first}: {len(t)} lines; header {t[:1]}; facility types "
                f"{sorted(set(l.split(',')[3] for l in t[1:] if l.count(',') >= 5))[:20] if v == '16' else ''}")
        elif r is not None:
            out(r.text[:300])


def ema_stats():
    out("\n===== EMA statistics-list =====")
    for pp in ["/content/corporate/language-masters/en/home/resources/statistics",
               "/content/corporate/language-masters/en/home/resources/statistics.html"]:
        r = get(f"{EMA}/bin/corporate-site/statistics-list", params={"searchString": "", "category": "",
                                                                     "pagePath": pp})
        if r is None or r.status_code != 200:
            continue
        try:
            items = r.json()
        except ValueError:
            out(r.text[:1000])
            continue
        out(f"  {len(items)} items")
        for it in items:
            dl = [(d.get("fileTitle"), d.get("fileLink")) for d in it.get("downloadInfo", [])]
            out(f"  - {it.get('title')!r} | {it.get('tags')} | {it.get('date')} | {dl[:4]}")
        break


if __name__ == "__main__":
    for f in (ema_stats, nems_year, ema_xls):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__} failed: {e!r}")
