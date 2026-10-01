"""
Singapore discovery, round 2 (after SINGAPORE_DISCOVERY.py found):
  - EMA 'Half-hourly System Demand Data' and 'Statistics' pages build their
    download lists in JavaScript (downloadItem.fileLink) - find the JSON /
    file endpoints they use.
  - EMA Singapore Energy Statistics workbooks (SES_tidy.xlsx, SES_tabular.xlsx)
    - list their sheets / series and whether any are monthly.
  - the rest of the data.gov.sg dataset listing (pages 200+).
  - SingStat TableBuilder searches for monthly electricity / gas tables.
  - EMC / NEMS price-and-demand pages (cert chain is incomplete: probe with verify=False).
"""
import io
import json
import re
import time
import warnings

import pandas as pd
import requests

warnings.filterwarnings("ignore")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json, text/html, */*"}
T = (10, 90)
EMA = "https://www.ema.gov.sg"


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
            time.sleep(10 * (i + 1))
            continue
        return r
    return r


def page_refs(url, verify=True):
    out(f"\n===== {url} =====")
    r = get(url, verify=verify)
    if r is None:
        return None
    t = r.text
    out(f"status {r.status_code} len {len(t)}")
    pats = [r'["\'](/[^"\']*\.(?:json|csv|xlsx?|zip|pdf)[^"\']*)["\']', r'["\'](https?://[^"\']+)["\']',
            r'data-[a-z-]+="([^"]{8,300})"', r'(/content/[^"\'\s<>]+)', r'(/bin/[^"\'\s<>]+)', r'(/api/[^"\'\s<>]+)']
    seen = set()
    for p in pats:
        for m in re.findall(p, t, re.I):
            if m in seen or re.search(r"\.(css|png|jpg|svg|woff2?|ico)(\?|$)|googletag|clientlibs", m, re.I):
                continue
            seen.add(m)
            out(f"  REF {m[:300]}")
    for kw in ["downloadItem", "fileLink", "ng-init", "ng-repeat", "angular", "fetch(", "$.ajax", "$.get", "XMLHttp"]:
        for mm in list(re.finditer(re.escape(kw), t))[:4]:
            out(f"  CTX[{kw}] ...{t[max(0, mm.start() - 400): mm.start() + 600]!r}")
    return r


def workbook(url, name):
    out(f"\n===== workbook {name} {url} =====")
    r = get(url)
    if r is None or r.status_code != 200:
        out(f"  status {getattr(r, 'status_code', None)}")
        return
    out(f"  {len(r.content)} bytes {r.headers.get('content-type')}")
    xl = pd.ExcelFile(io.BytesIO(r.content))
    out(f"  sheets ({len(xl.sheet_names)}): {xl.sheet_names}")
    for s in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=s, header=None, nrows=400)
        out(f"\n  --- sheet {s!r} shape(first 400 rows) {df.shape}")
        out(df.head(12).to_string(max_colwidth=40, max_cols=14)[:2500])
        if "tidy" in name.lower():
            full = pd.read_excel(xl, sheet_name=s)
            out(f"  full shape {full.shape}; columns {list(full.columns)}")
            for c in full.columns:
                if full[c].dtype == object and full[c].nunique() < 200:
                    out(f"    {c}: {sorted(map(str, full[c].dropna().unique()))[:80]}")
                elif re.search(r"year|month|date|period", str(c), re.I):
                    out(f"    {c}: min {full[c].min()} max {full[c].max()} sample {list(full[c].dropna().unique()[:12])}")


def dgs_rest():
    out("\n===== data.gov.sg datasets pages 190+ =====")
    base = "https://api-production.data.gov.sg/v2/public/api"
    key = re.compile(r"electric|natural gas|town gas|piped gas|\bgas\b|energy|power|system demand|lng|generation", re.I)
    page, pages = 190, 999
    while page <= pages:
        r = get(f"{base}/datasets", params={"page": page})
        if r is None or r.status_code != 200:
            out(f"page {page}: {getattr(r, 'status_code', None)}")
            time.sleep(5)
            page += 1
            continue
        d = r.json().get("data", {})
        pages = d.get("pages", page) or page
        for s in d.get("datasets", []):
            nm = s.get("name", "")
            ag = str(s.get("managedByAgencyName", ""))
            if key.search(nm) or "Energy Market" in ag:
                out(f"DS {s.get('datasetId')} | {nm} | {ag[:40]} | cov={s.get('coverageStart')}..{s.get('coverageEnd')} "
                    f"| upd={s.get('lastUpdatedAt')}")
        page += 1
        time.sleep(0.6)
    out(f"done through page {page - 1}/{pages}")


def singstat():
    out("\n===== SingStat searches =====")
    for kw in ["Electricity Generation", "Electricity Consumption", "Piped Gas", "Natural Gas Consumption",
               "System Demand", "Peak", "Gas Sales", "Electricity, Gas"]:
        r = get("https://tablebuilder.singstat.gov.sg/api/table/resourceid",
                params={"keyword": kw, "searchOption": "title"})
        try:
            data = r.json().get("Data")
            recs = data.get("records", []) if isinstance(data, dict) else (data or [])
        except Exception as e:
            out(f"[{kw}] {e}")
            continue
        out(f"[{kw}] {len(recs)}")
        for t in recs[:40]:
            out(f"  {t.get('id')} | {t.get('title')}")


def singstat_table(tid):
    r = get(f"https://tablebuilder.singstat.gov.sg/api/table/tabledata/{tid}", params={"limit": 3000})
    try:
        d = r.json().get("Data") or {}
    except Exception as e:
        out(f"{tid}: {e} {getattr(r, 'status_code', None)}")
        return
    out(f"\nTABLE {tid}: {d.get('title')} | freq {d.get('frequency')} | dataLastUpdated {d.get('dataLastUpdated')}")
    for row in d.get("row", [])[:30]:
        cols = row.get("columns") or []
        out(f"  {row.get('seriesNo')} {row.get('rowText')!r} [{row.get('uoM')}] n={len(cols)} "
            f"first={cols[:1]} last={cols[-2:]}")


if __name__ == "__main__":
    steps = [
        lambda: page_refs(f"{EMA}/resources/statistics/half-hourly-system-demand-data"),
        lambda: page_refs(f"{EMA}/resources/statistics"),
        lambda: workbook(f"{EMA}/content/dam/corporate/resources/singapore-energy-statistics/excel/SES_tidy.xlsx.coredownload.xlsx", "SES_tidy"),
        lambda: workbook(f"{EMA}/content/dam/corporate/resources/singapore-energy-statistics/excel/SES_tabular.xlsx.coredownload.xlsx", "SES_tabular"),
        singstat,
        lambda: [singstat_table(t) for t in ["M890381", "M890371", "M891621", "M891111", "M891291"]],
        lambda: page_refs("https://www.emcsg.com/marketdata/priceinformation", verify=False),
        lambda: page_refs("https://www.nems.emcsg.com/nems-prices"),
        dgs_rest,
    ]
    for f in steps:
        try:
            f()
        except Exception as e:
            out(f"!! step failed: {e!r}")
