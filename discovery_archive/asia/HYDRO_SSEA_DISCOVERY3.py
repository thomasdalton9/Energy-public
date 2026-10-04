"""
South & Southeast Asia hydro reservoir discovery, round 3 (after HYDRO_SSEA_DISCOVERY2.py):
  Pakistan  WAPDA media library (wp-json/wp/v2/media search 'Daily-DataPress' / 'GRAPH'): is there a dated archive
            of the daily water-data xls?; the 2024 xls content
  Sri Lanka PUCSL storage-rainfall: earliest year, reservoir names; /api/reservoir/inflow with dateAggregation
  Vietnam   vndms.gov.vn JS: reservoir (ho chua) endpoints
  Laos      MRC data portal (time-series API) for Nam Ngum / Nam Theun reservoirs
"""
import io
import re
from datetime import date, timedelta

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "text/html,application/json,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}
T = (15, 60)
s = requests.Session()
s.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    try:
        r = s.get(u, timeout=k.pop("timeout", T), verify=False, **k)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {type(e).__name__}: {str(e)[:200]}")
        return None


def wapda():
    out("\n=========== WAPDA media")
    seen = []
    for q in ("Daily-Data", "DataPress", "Press-Release", "GRAPH-DG", "water situation", "river flow", "Daily Data"):
        for page in (1, 2):
            r = get("https://wapda.gov.pk/wp-json/wp/v2/media", params={"search": q, "per_page": 100, "page": page,
                                                                          "_fields": "id,date,source_url,title"})
            if r is None or not r.ok:
                break
            try:
                j = r.json()
            except ValueError:
                break
            out(f"  q={q} p{page} total={r.headers.get('X-WP-Total')}: {len(j)}")
            for it in j:
                if it["source_url"] not in seen:
                    seen.append(it["source_url"])
                    out(f"    {it['date']} {it['source_url']}")
            if len(j) < 100:
                break
    # newest media of xls type
    r = get("https://wapda.gov.pk/wp-json/wp/v2/media", params={"mime_type": "application/vnd.ms-excel", "per_page": 50,
                                                                  "_fields": "id,date,source_url"})
    if r is not None and r.ok:
        for it in r.json():
            out(f"    xls {it['date']} {it['source_url']}")
    r = get("https://wapda.gov.pk/wp-json/wp/v2/posts", params={"categories_slug": "river-data", "per_page": 20,
                                                                  "_fields": "id,date,modified,slug,link"})
    r = get("https://wapda.gov.pk/wp-json/wp/v2/categories", params={"search": "river", "_fields": "id,name,slug,count"})
    if r is not None and r.ok:
        out("   cats:", r.text[:500])
        for c in r.json():
            p = get("https://wapda.gov.pk/wp-json/wp/v2/posts", params={"categories": c["id"], "per_page": 50,
                                                                          "_fields": "id,date,modified,slug,link"})
            if p is not None and p.ok:
                out("   posts:", p.text[:3000])
    for u in ("https://wapda.gov.pk/wp-content/uploads/2024/12/Daily-DataPress-Release-19-12-2024.xls",
              "https://wapda.gov.pk/wp-content/uploads/2024/12/GRAPH-DG-16-for-MAIL-2.xls"):
        r = get(u)
        if r is not None and r.ok:
            try:
                for sh, d in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
                    out(f"   sheet {sh} {d.shape}\n{d.head(40).to_string()[:4000]}")
            except Exception as e:  # noqa: BLE001
                out("   xls err", e, r.content[:200])


def pucsl():
    out("\n=========== PUCSL")
    A = "https://gendata.pucsl.gov.lk/api/"
    hd = {"Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"}
    for y in (2005, 2008, 2010, 2012, 2013, 2014):
        r = get(A + "reservoir/storage-rainfall", params={"dateAggregation": "day", "from": f"{y}-01-01T00:00:00.000Z",
                                                          "to": f"{y}-01-05T00:00:00.000Z"}, headers=hd)
        if r is not None:
            out("   ", r.text[:300])
    r = get(A + "reservoir/storage-rainfall", params={"dateAggregation": "day", "from": "2025-01-01T00:00:00.000Z",
                                                      "to": "2026-01-01T00:00:00.000Z"}, headers=hd, timeout=(15, 180))
    if r is not None and r.ok:
        j = r.json()["data"]
        d = pd.DataFrame(j)
        out(f"   one year: {len(d)} rows; names {d.reservoirName.value_counts().to_dict()}")
        out(d.groupby("reservoirName")[["molBelowSpillInM", "storageInGwh"]].describe().to_string())
        tot = d.groupby("reportDate").storageInGwh.sum()
        out("   total GWh: ", tot.describe().to_string())
    for agg in ("day", "month"):
        r = get(A + "reservoir/inflow", params={"dateAggregation": agg, "from": "2026-08-01T00:00:00.000Z",
                                                "to": "2026-10-03T00:00:00.000Z"}, headers=hd)
        if r is not None:
            out("   ", r.text[:800])
    for p in ("metadata/power-plant-complexes",):
        r = get(A + p, headers=hd)
        if r is not None:
            out("   ", r.text[:1500])


def vndms():
    out("\n=========== VNDMS")
    r = get("https://vndms.gov.vn/")
    if r is None:
        return
    h = r.text
    for m in list(re.finditer(r"(?i)(h[oồ]\s*ch[uứ]a|hochua|reservoir|thuydien|thuy-dien|hồ thủy điện)", h))[:20]:
        out("  html: " + h[max(0, m.start() - 200):m.start() + 300].replace("\n", " "))
    for js in re.findall(r'<script[^>]+src=["\']([^"\']+vndms[^"\']+)', h):
        k = get(js)
        if k is None or not k.ok:
            continue
        urls = sorted(set(re.findall(r'["\'`](/[A-Za-z][\w\-]*(?:/[\w\-{}.]+)+)["\'`]', k.text)))
        out(f"   paths in {js[-40:]} ({len(urls)}):", [u for u in urls if not u.endswith((".png", ".svg", ".css"))][:150])
        for m in list(re.finditer(r"(?i)(hochua|ho_chua|HoChua|reservoir|thuydien)", k.text))[:15]:
            out("    js: " + k.text[max(0, m.start() - 200):m.start() + 250].replace("\n", " "))


def mrc():
    out("\n=========== MRC (Laos)")
    for u in ("https://portal.mrcmekong.org/time-series/water-level", "https://api.mrcmekong.org/api/v1/ts/highcharts/stations",
              "https://portal.mrcmekong.org/api/v1/ts/inventory/stations"):
        r = get(u)
        if r is not None:
            out("   ", r.text[:800])


for f in (wapda, pucsl, vndms, mrc):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
