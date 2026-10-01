"""
One-off discovery for Singapore power and gas demand sources (runs in
GitHub Actions; the sites are blocked from the editing sandbox).

Probes:
  1. EMA statistics pages (half-hourly system demand, energy statistics) -
     prints every data/download link.
  2. SingStat TableBuilder API - tables matching electricity / gas keywords.
  3. EMC (NEMS) market data pages - prints download links.
  4. data.gov.sg v2 API - collections and datasets whose names mention
     electricity / gas / energy / demand, plus their metadata and a few rows.

Findings (rounds 1-7, SINGAPORE_DISCOVERY2..7.py), used by asia/SINGAPORE_POWER.py and SINGAPORE_GAS.py:
  - EMA half-hourly system demand: weekly .xls files 2014->, listed by
    /bin/corporate-site/half-hourly-list (System Demand actual, NEM demand actual + forecast, MW).
  - EMC NEMS data download: /api/sitecore/DataSync/DataDownload(ByYear)?value=16 = metered generation by
    facility type (half-hourly MWh, 2021->, five-year rolling); value=1 = USEP price + demand forecast.
  - SingStat M890831 Electricity Generation, Monthly (GWh, 1975->); M890371 Piped (town) gas sales, quarterly.
  - EMA SES tidy workbook (annual): T1.1 NG imports pipeline/LNG, T2.1 gas into power, T2.2 fuel mix,
    T3.2 electricity consumption by sector, T3.7 NG final consumption by sector.
  - data.gov.sg EMA datasets are old snapshots (mostly to 2020/2021); no monthly gas consumption or imports
    anywhere; SingStat T010002 detailed trade table has no working tabledata API (HTTP 400).
"""
import json
import re
import time

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json, text/html, */*"}
T = (10, 60)
KEY = re.compile(r"electric|natural gas|town gas|\bgas\b|energy|power|demand|lng|generation", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    r = None
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T, **kw)
        except requests.RequestException as e:
            out(f"  ERR {url}: {e}")
            return None
        if r.status_code == 429:
            out(f"  429 on {url}; sleeping")
            time.sleep(10 * (i + 1))
            continue
        return r
    return r


def dgs():
    out("\n===== data.gov.sg collections =====")
    base = "https://api-production.data.gov.sg/v2/public/api"
    page, pages = 1, 1
    while page <= pages and page <= 60:
        r = get(f"{base}/collections", params={"page": page})
        if r is None or r.status_code != 200:
            out(f"collections page {page}: {r.status_code if r is not None else 'none'} "
                f"{r.text[:300] if r is not None else ''}")
            break
        j = r.json()
        if page == 1:
            out(json.dumps(j)[:1500])
        d = j.get("data", {})
        pages = d.get("pages", d.get("totalPages", 1)) or 1
        for c in d.get("collections", []):
            nm = c.get("name", "")
            ag = json.dumps(c.get("managedBy", c.get("managedByAgencyName", "")))
            if KEY.search(nm) or "Energy Market" in ag:
                out(f"COLL {c.get('collectionId')} | {nm} | {ag[:80]} | {c.get('lastUpdatedAt')} | "
                    f"datasets={c.get('childDatasets')}")
        page += 1
        time.sleep(1)
    out(f"collections pages scanned: {page - 1} / {pages}")

    out("\n===== data.gov.sg datasets =====")
    ds_hits = []
    page, pages = 1, 1
    while page <= pages and page <= 200:
        r = get(f"{base}/datasets", params={"page": page})
        if r is None or r.status_code != 200:
            out(f"datasets page {page}: {r.status_code if r is not None else 'none'} "
                f"{r.text[:300] if r is not None else ''}")
            break
        j = r.json()
        if page == 1:
            out(json.dumps(j)[:1500])
        d = j.get("data", {})
        pages = d.get("pages", d.get("totalPages", 1)) or 1
        for s in d.get("datasets", []):
            nm = s.get("name", "")
            ag = json.dumps(s.get("managedByAgencyName", s.get("managedBy", "")))
            if KEY.search(nm) or "Energy Market" in ag:
                ds_hits.append(s)
                out(f"DS {s.get('datasetId')} | {nm} | {ag[:60]} | fmt={s.get('format')} | "
                    f"cov={s.get('coverageStart')}..{s.get('coverageEnd')} | upd={s.get('lastUpdatedAt')}")
        page += 1
        time.sleep(0.7)
    out(f"dataset pages scanned: {page - 1} / {pages}")

    want = re.compile(r"half.?hourly|system demand|electricity generation|electricity consumption|natural gas|"
                      r"town gas|peak", re.I)
    for s in ds_hits:
        if not want.search(s.get("name", "")):
            continue
        did = s.get("datasetId")
        time.sleep(2)
        r = get("https://data.gov.sg/api/action/datastore_search", params={"resource_id": did, "limit": 3,
                                                                          "sort": "_id desc"})
        if r is None:
            continue
        try:
            res = r.json().get("result", {})
            out(f"\nROWS {did} {s.get('name')}: total={res.get('total')} fields="
                f"{[f.get('id') for f in res.get('fields', [])]}")
            for rec in res.get("records", [])[:3]:
                out(f"   {rec}")
        except ValueError:
            out(f"ROWS {did}: {r.status_code} {r.text[:200]}")


def links(url, pat=r"\.(csv|xlsx?|pdf|zip)|download|statistic|demand|api", limit=200):
    out(f"\n===== links on {url} =====")
    r = get(url)
    if r is None:
        return ""
    out(f"status {r.status_code} len {len(r.text)} final {r.url}")
    seen = set()
    for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", txt).strip()
        if (re.search(pat, href, re.I) or re.search(pat, t, re.I)) and href not in seen:
            seen.add(href)
            out(f"  {t[:80]!r} -> {href}")
            if len(seen) >= limit:
                break
    for m in sorted(set(re.findall(r'https?://[^"\'\s<>]+\.(?:csv|xlsx?|zip)', r.text)))[:50]:
        out(f"  RAW {m}")
    return r.text


def ema():
    for u in ["https://www.ema.gov.sg/resources/statistics/half-hourly-system-demand-data",
              "https://www.ema.gov.sg/resources/statistics",
              "https://www.ema.gov.sg/resources/singapore-energy-statistics",
              "https://www.ema.gov.sg/resources/singapore-energy-statistics/chapter3",
              "https://www.ema.gov.sg/resources/singapore-energy-statistics/chapter2"]:
        txt = links(u)
        if txt and "half-hourly" in u:
            i = txt.lower().find("system demand")
            out(txt[max(0, i - 500): i + 3000] if i >= 0 else txt[:3000])


def singstat():
    out("\n===== SingStat TableBuilder =====")
    for kw in ["electricity", "natural gas", "town gas", "energy", "LNG"]:
        r = get("https://tablebuilder.singstat.gov.sg/api/table/resourceid",
                params={"keyword": kw, "searchOption": "all"})
        if r is None:
            continue
        try:
            j = r.json()
        except ValueError:
            out(f"{kw}: {r.status_code} {r.text[:300]}")
            continue
        data = j.get("Data")
        recs = data.get("records", []) if isinstance(data, dict) else (data or [])
        out(f"\n[{kw}] {len(recs)} tables")
        for t in recs[:60]:
            out(f"  {t.get('id')} | {t.get('title')} | {t.get('frequency', '')}")
        if not recs:
            out(json.dumps(j)[:800])


def emc():
    for u in ["https://www.emcsg.com/marketdata/priceinformation",
              "https://www.emcsg.com/marketdata",
              "https://www.nems.emcsg.com/nems-prices"]:
        links(u, pat=r"\.(csv|xlsx?)|download|demand|price|export")


if __name__ == "__main__":
    for f in (ema, singstat, emc, dgs):
        try:
            f()
        except Exception as e:  # keep probing the rest
            out(f"!! {f.__name__} failed: {e!r}")
