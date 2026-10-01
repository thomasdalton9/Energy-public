"""
Singapore discovery, round 3. Round 2 found:
  - EMA half-hourly page references
      /content/dam/corporate/statistics/half-hourly-data.csv  and a
      ...half-hourly-system-demand-data.searchresults.json/... AEM search endpoint
  - SingStat M890831 'Electricity Generation, Monthly' (data.gov.sg d_ae4afbaf5bc96bde19d8ce85810ab9f4)
  - EMA /resources/statistics page is a Vue app (item.downloadInfo) - endpoint unknown
  - NEMS prices page embeds [USEP, Demand, Solar, VCP] charts
This round fetches those endpoints and prints their structure.
"""
import io
import re
import time
import warnings

import pandas as pd
import requests

warnings.filterwarnings("ignore")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json, text/html, */*"}
T = (10, 120)
EMA = "https://www.ema.gov.sg"


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
        return r
    except requests.RequestException as e:
        out(f"ERR {url}: {e}")
        return None


def show_csv(content, label):
    text = content.decode("utf-8-sig", errors="replace")
    lines = text.splitlines()
    out(f"  {label}: {len(lines)} lines")
    for l in lines[:15]:
        out(f"   | {l[:250]}")
    out("   ...")
    for l in lines[-8:]:
        out(f"   | {l[:250]}")


def ctx(text, kws, width=700, n=4):
    for kw in kws:
        for m in list(re.finditer(re.escape(kw), text, re.I))[:n]:
            out(f"  CTX[{kw}] {text[max(0, m.start() - width // 3): m.start() + width]!r}")


def half_hourly():
    out("\n===== EMA half-hourly =====")
    r = get(f"{EMA}/content/dam/corporate/statistics/half-hourly-data.csv")
    if r is not None and r.status_code == 200:
        show_csv(r.content, "half-hourly-data.csv")
    base = (f"{EMA}/content/corporate/language-masters/en/home/resources/statistics/"
            "half-hourly-system-demand-data.searchresults.json/_jcr_content/root/container_630399545/"
            "container_464942973/search")
    for q in ["", "?fulltext=&orderby=@jcr:content/jcr:lastModified&orderby.sort=desc", "?p.limit=50",
              "?q=&offset=0&limit=50"]:
        r = get(base + q)
        if r is not None:
            out(f"  {r.text[:3000]}")
    r = get(f"{EMA}/resources/statistics/half-hourly-system-demand-data")
    if r is not None:
        ctx(r.text, ["searchresults", "half-hourly-data", "itemTemplate", "data-search", "loadingIndicator"], 1500, 3)
    for u in [f"{EMA}/content/dam/corporate/resources/statistics/half-hourly-data.json",
              f"{EMA}/content/dam/corporate/resources/statistics/half-hourly-data.1.json",
              f"{EMA}/content/dam/corporate/resources/statistics/half-hourly-data.infinity.json"]:
        r = get(u)
        if r is not None:
            out(f"  {r.text[:1500]}")


def statistics_page():
    out("\n===== EMA statistics page =====")
    r = get(f"{EMA}/resources/statistics")
    if r is None:
        return
    t = r.text
    ctx(t, ["new Vue", "axios", "fetch(", ".json", "servlet", "/bin/", "api/", "endpoint", "data-url", "dataUrl",
            "downloadInfo", "statisticsList", "mounted"], 900, 3)
    for js in sorted(set(re.findall(r'src="(/etc\.clientlibs/[^"]+\.js)"', t))):
        jr = get(EMA + js)
        if jr is None:
            continue
        for kw in ["statistic", "downloadInfo", "searchresults", ".json"]:
            for m in list(re.finditer(re.escape(kw), jr.text))[:4]:
                out(f"  JS {js[-50:]} [{kw}] {jr.text[max(0, m.start() - 300): m.start() + 500]!r}")


def singstat_monthly():
    out("\n===== SingStat M890831 Electricity Generation, Monthly =====")
    r = get("https://tablebuilder.singstat.gov.sg/api/table/tabledata/M890831", params={"limit": 3000})
    d = (r.json().get("Data") or {}) if r is not None else {}
    out(f"title {d.get('title')} freq {d.get('frequency')} updated {d.get('dataLastUpdated')} "
        f"footnote {str(d.get('footnote'))[:600]}")
    for row in d.get("row", []):
        cols = row.get("columns") or []
        out(f"  {row.get('seriesNo')} {row.get('rowText')!r} [{row.get('uoM')}] n={len(cols)} first={cols[:1]} "
            f"last={cols[-3:]}")
    r = get("https://data.gov.sg/api/action/datastore_search",
            params={"resource_id": "d_ae4afbaf5bc96bde19d8ce85810ab9f4", "limit": 20})
    if r is not None:
        res = r.json().get("result", {})
        out(f"  data.gov.sg rows total={res.get('total')} fields={[f.get('id') for f in res.get('fields', [])][:8]}...")
        for rec in res.get("records", []):
            out(f"   {rec.get('DataSeries')!r}: {[(k, rec[k]) for k in list(rec)[:4]]}")
    # other EMA-sourced monthly SingStat tables (M8908xx / M8903xx)
    for tid in ["M890821", "M890811", "M890851", "M890861", "M890871", "M890361", "M890391", "M890351"]:
        r = get(f"https://tablebuilder.singstat.gov.sg/api/table/tabledata/{tid}", params={"limit": 5})
        try:
            dd = r.json().get("Data") or {}
            out(f"  {tid}: {dd.get('title')} | {dd.get('frequency')} | rows "
                f"{[x.get('rowText') for x in dd.get('row', [])][:10]}")
        except Exception as e:
            out(f"  {tid}: {e}")
        time.sleep(0.5)


def nems():
    out("\n===== NEMS prices page =====")
    r = get("https://www.nems.emcsg.com/nems-prices")
    if r is None:
        return
    t = r.text
    ctx(t, ["ajax", "fetch(", "api/", "/umbraco", "surface", ".json", "Demand", "csv", "download", "Export"], 900, 3)


if __name__ == "__main__":
    for f in (half_hourly, singstat_monthly, statistics_page, nems):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__} failed: {e!r}")
