"""
Discovery probe: HISTORY of Egypt natural gas data (production, consumption by sector, pipeline imports
from Israel, LNG imports/exports, pipeline exports to Jordan) automatable from GitHub Actions (Oct 2026).

Starts from EGYPT_ISRAEL_GAS_POWER_PROBE.py (CAPMAS energy balance annual 2014-2022, MOEE fuel for power annual,
EGAS / petroleum.gov.eg no statistics pages).

RUN findings: see bottom of this docstring (filled in after the Actions runs).
"""

import io
import json
import os
import re
import sys
import time
import zipfile
from urllib.parse import urljoin, quote

import requests

try:
    from curl_cffi import requests as creq
except Exception:  # noqa: BLE001
    creq = None

H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}
T = (15, 60)
ROUND = os.environ.get("ROUND", "1")
KEY = re.compile(r"(stat|report|annual|month|bulletin|indicator|xls|csv|pdf|gas|product|export|import|lng|"
                 r"إحصا|تقرير|نشر|غاز|انتاج|إنتاج|مؤشر)", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, cffi_fallback=True, quiet=False, **kw):
    """requests first, then curl_cffi (chrome) on error / 403 / 503."""
    r = None
    try:
        r = requests.get(url, headers=H, timeout=T, allow_redirects=True, **kw)
        if r.status_code not in (403, 503, 429) or not cffi_fallback or creq is None:
            return r
        if not quiet:
            out(f"  requests {r.status_code} -> curl_cffi")
    except requests.RequestException as e:
        if not quiet:
            out(f"  requests ERROR {type(e).__name__}: {str(e)[:120]}")
        if not cffi_fallback or creq is None:
            return None
    try:
        return creq.get(url, impersonate="chrome", timeout=60, allow_redirects=True, **kw)
    except Exception as e:  # noqa: BLE001
        if not quiet:
            out(f"  curl_cffi ERROR {type(e).__name__}: {str(e)[:120]}")
        return r


def page(url, max_links=50, pat=KEY, show=True):
    out(f"\n=== {url}")
    r = get(url)
    if r is None:
        return None
    ct = r.headers.get("content-type", "")
    m = re.search(r"<title[^>]*>(.*?)</title>", r.text[:30000], re.S | re.I) if "html" in ct else None
    title = " ".join(m.group(1).split())[:90] if m else ""
    out(f"  {r.status_code} {ct[:35]} {len(r.content)}B final={str(r.url)[:110]} title={title!r}")
    if r.status_code != 200 or "html" not in ct or not show:
        if r.status_code != 200:
            out("  body:", " ".join(r.text[:200].split()))
        return r
    seen, n = set(), 0
    for href, txt in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = " ".join(re.sub(r"<[^>]+>", " ", txt).split())[:70]
        full = urljoin(str(r.url), href)
        if full in seen or not (pat.search(full) or pat.search(t)):
            continue
        seen.add(full)
        out(f"   - {full[:160]} | {t}")
        n += 1
        if n >= max_links:
            out("   ...")
            break
    files = sorted(set(re.findall(r'["\']([^"\'\s]+\.(?:pdf|xlsx?|csv|zip))["\']', r.text, re.I)))
    if files:
        out(f"  file refs ({len(files)}):", [f[:120] for f in files[:40]])
    return r


def cdx(pattern, limit=300, filt=None):
    url = ("http://web.archive.org/cdx/search/cdx?url=" + quote(pattern, safe="*/:")
           + f"&output=json&collapse=urlkey&limit={limit}")
    if filt:
        url += "&filter=" + quote(filt)
    out(f"\n=== CDX {pattern} {filt or ''}")
    for att in range(3):
        try:
            r = requests.get(url, headers=H, timeout=(20, 120))
            if r.status_code in (429, 503):
                out("  ", r.status_code, "retry")
                time.sleep(15 * (att + 1))
                continue
            rows = r.json()[1:] if r.text.strip().startswith("[") else []
            out(f"  {r.status_code} {len(rows)} rows")
            return rows
        except Exception as e:  # noqa: BLE001
            out("  CDX ERROR", type(e).__name__, str(e)[:120])
            time.sleep(10)
    return []


# ------------------------------------------------------------------ JODI
JODI_ZIP = "https://www.jodidata.org/_resources/files/downloads/gas-data/GAS_world_NewFormat.zip"


def jodi():
    out("\n################ JODI Gas world NewFormat zip")
    import pandas as pd
    r = get(JODI_ZIP)
    if r is None:
        return
    out("  ", r.status_code, len(r.content), r.headers.get("content-type"), r.headers.get("last-modified"))
    if r.status_code != 200 or len(r.content) < 10000:
        return
    z = zipfile.ZipFile(io.BytesIO(r.content))
    out("  members:", z.namelist())
    frames = []
    for n in z.namelist():
        if n.lower().endswith(".csv"):
            frames.append(pd.read_csv(z.open(n), dtype=str))
    d = pd.concat(frames, ignore_index=True)
    out("  columns:", list(d.columns), d.shape)
    eg = d[d["REF_AREA"].str.upper() == "EG"]
    out("  Egypt rows:", eg.shape, "TIME", eg["TIME_PERIOD"].min(), "..", eg["TIME_PERIOD"].max())
    out("  Egypt flows:", sorted(eg["FLOW_BREAKDOWN"].unique()))
    out("  units:", sorted(eg["UNIT_MEASURE"].unique()), "assessment:", sorted(eg.get("ASSESSMENT_CODE", pd.Series()).dropna().unique()))
    e = eg[eg["UNIT_MEASURE"] == "M3"].copy()
    e["v"] = pd.to_numeric(e["OBS_VALUE"], errors="coerce")
    e["y"] = e["TIME_PERIOD"].str[:4]
    piv = e.pivot_table(index="y", columns="FLOW_BREAKDOWN", values="v", aggfunc="sum") / 1000.0
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 40)
    out("  Egypt annual sums, bcm (M3 unit is million m3):")
    out(piv.round(2).to_string())
    for fl in sorted(e["FLOW_BREAKDOWN"].unique()):
        s = e[(e["FLOW_BREAKDOWN"] == fl) & (e["v"].abs() > 0)]
        if len(s):
            out(f"   {fl:10s} nonzero {s['TIME_PERIOD'].min()}..{s['TIME_PERIOD'].max()} n={len(s)} last={s.sort_values('TIME_PERIOD')['v'].iloc[-1]}")
    tail = e[e["TIME_PERIOD"] >= "2025-07"].pivot_table(index="TIME_PERIOD", columns="FLOW_BREAKDOWN", values="v")
    out(tail.round(0).to_string())


# ------------------------------------------------------------------ Comtrade
CT = "https://comtradeapi.un.org/public/v1/preview/C/{freq}/HS"


def comtrade(freq, **params):
    for attempt in range(4):
        try:
            r = requests.get(CT.format(freq=freq), params=params, headers=H, timeout=T)
        except Exception as e:  # noqa: BLE001
            out("   ERR", type(e).__name__, str(e)[:100])
            time.sleep(5)
            continue
        if r.status_code == 429:
            time.sleep(8 * (attempt + 1))
            continue
        time.sleep(1.3)
        if r.status_code != 200:
            out("   HTTP", r.status_code, r.text[:200])
            return None
        j = r.json()
        if j.get("error"):
            out("   error:", j.get("error"))
        return j.get("data") or []
    return None


def ct_summary(rows, label):
    agg = {}
    for x in rows or []:
        if x.get("partnerCode") not in (0,) and label.endswith("world"):
            continue
        k = (x.get("flowCode"), x.get("cmdCode"), x.get("partnerDesc") or x.get("partnerCode"))
        a = agg.setdefault(k, [0.0, 0.0, 0.0])
        a[0] += x.get("netWgt") or 0
        a[1] += x.get("primaryValue") or 0
        a[2] += x.get("qty") or 0
    for k, v in sorted(agg.items()):
        out(f"    {label} {k} netWgt_t={v[0]/1000:,.0f} USD_m={v[1]/1e6:,.1f} qty={v[2]:,.0f}")


def comtrade_probe():
    out("\n################ UN Comtrade public preview")
    for y in range(2014, 2026):
        rows = comtrade("A", reporterCode=818, period=y, cmdCode="271111,271121", flowCode="M,X")
        out(f"  Egypt annual {y}: {None if rows is None else len(rows)} rows")
        if rows:
            ct_summary([x for x in rows if x.get("partnerCode") == 0], "world")
            big = sorted([x for x in rows if x.get("partnerCode") != 0], key=lambda x: -(x.get("primaryValue") or 0))[:5]
            for x in big:
                out(f"     top {x.get('flowCode')} {x.get('cmdCode')} {x.get('partnerDesc')} t={(x.get('netWgt') or 0)/1000:,.0f} USD_m={(x.get('primaryValue') or 0)/1e6:,.1f}")
    out("  -- Egypt monthly availability")
    for p in ["201601", "201901", "202101", "202201", "202301", "202306", "202401", "202406", "202412",
              "202501", "202506", "202512", "202601", "202604"]:
        rows = comtrade("M", reporterCode=818, period=p, cmdCode="2711,271111,271121", flowCode="M,X")
        out(f"  Egypt monthly {p}: {None if rows is None else len(rows)} rows")
        if rows:
            ct_summary([x for x in rows if x.get("partnerCode") == 0], "world")
    out("  -- mirrors (annual): Israel X to Egypt, Jordan M from Egypt, world LNG exporters to Egypt")
    for y in [2019, 2020, 2021, 2022, 2023, 2024, 2025]:
        for rep, fl in [(376, "X"), (400, "M"), (842, "X")]:
            rows = comtrade("A", reporterCode=rep, period=y, cmdCode="271111,271121", flowCode=fl, partnerCode=818)
            out(f"  {rep} {fl} {y}: {None if rows is None else len(rows)}")
            ct_summary(rows, f"r{rep}")
    out("  -- mirror monthly Israel exports to Egypt")
    for p in ["202001", "202201", "202401", "202501", "202506"]:
        rows = comtrade("M", reporterCode=376, period=p, cmdCode="271121,271111", flowCode="X", partnerCode=818)
        out(f"  376 X {p}: {None if rows is None else len(rows)}")
        ct_summary(rows, "r376")


# ------------------------------------------------------------------ CBE
def cbe():
    out("\n################ Central Bank of Egypt")
    for u in ["https://www.cbe.org.eg/en/economic-research/statistics",
              "https://www.cbe.org.eg/en/economic-research/publications",
              "https://www.cbe.org.eg/en/economic-research/publications/monthly-statistical-bulletin",
              "https://www.cbe.org.eg/en/economic-research/economic-reports/monthly-statistical-bulletin",
              "https://www.cbe.org.eg/en/economic-research/time-series",
              "https://www.cbe.org.eg/en/economic-research/statistics/external-sector",
              "https://www.cbe.org.eg/en/economic-research/statistics/real-sector",
              "https://www.cbe.org.eg/ar/economic-research/statistics"]:
        page(u, max_links=40, pat=re.compile(r"(bulletin|statistic|time-series|real|petrol|xls|pdf|external|trade|review|report)", re.I))


# ------------------------------------------------------------------ EGAS / EGPC / Ministry
def sites():
    out("\n################ EGAS / EGPC / Ministry of Petroleum / IDSC")
    for u in ["https://www.egas.com.eg/", "https://www.egas.com.eg/ar", "https://egas.com.eg/web/guest/home",
              "https://www.egas.com.eg/web/guest/annual-report", "https://www.egas.com.eg/web/guest/statistics",
              "https://www.egpc.com.eg/", "http://www.egpc.com.eg/", "https://egpc.com.eg/ar",
              "https://www.petroleum.gov.eg/", "https://www.petroleum.gov.eg/ar/", "https://www.petroleum.gov.eg/en/",
              "https://www.petroleum.gov.eg/en/gas-and-petrochemicals/Pages/gas.aspx",
              "https://www.petroleum.gov.eg/ar/MediaCenter/Pages/default.aspx",
              "https://www.idsc.gov.eg/", "https://www.mped.gov.eg/", "https://ganope.com/", "https://www.ganope.com/",
              "https://www.echem.com.eg/", "https://www.egyptoil-gas.com/"]:
        page(u, max_links=35)


def wayback():
    out("\n################ Wayback CDX")
    for pat_, f in [("egas.com.eg/*", "mimetype:application/pdf"),
                    ("egas.com.eg/*", "urlkey:.*(statist|report|annual|indicat|product).*"),
                    ("www.egpc.com.eg/*", "mimetype:application/pdf"),
                    ("egpc.com.eg/*", "urlkey:.*(statist|report|annual|product).*"),
                    ("petroleum.gov.eg/*", "mimetype:application/pdf"),
                    ("www.petroleum.gov.eg/*", "urlkey:.*(statist|indicat|report|achiev|annual).*"),
                    ("cbe.org.eg/-/media/*", "urlkey:.*(bulletin|statistic).*")]:
        rows = cdx(pat_, limit=400, filt=f)
        for row in rows[:60]:
            out("   ", row[1], row[3][:25] if len(row) > 3 else "", row[2][:170])
        time.sleep(3)


# ------------------------------------------------------------------ CAPMAS
def capmas():
    out("\n################ CAPMAS subjects / indicators with month data")
    api = "https://www.capmas.gov.eg:8080/api/"
    r = get(api + "Subject")
    if r is None or r.status_code != 200:
        out("  Subject failed", None if r is None else r.status_code)
        return
    try:
        js = r.json()
    except Exception:  # noqa: BLE001
        out("  not json", r.text[:300])
        return
    data = js.get("data") if isinstance(js, dict) else js
    out("  Subject type", type(data).__name__, len(data) if hasattr(data, "__len__") else "")
    subs = []

    def walk(o, depth=0):
        if isinstance(o, dict):
            name = o.get("name") or o.get("title")
            tr = o.get("subjectTranslations") or o.get("subSubjectTranslations") or o.get("translations") or []
            en = next((t.get("name") or t.get("title") for t in tr if isinstance(t, dict) and t.get("locale") == "en"), "")
            if o.get("id") is not None and name:
                subs.append((depth, o.get("id"), name, en, [k for k in o.keys()][:12]))
            for v in o.values():
                walk(v, depth + 1)
        elif isinstance(o, list):
            for v in o:
                walk(v, depth + 1)
    walk(data)
    for s in subs[:200]:
        out("   ", s[0], s[1], s[2][:50], "|", (s[3] or "")[:60])
    pat = re.compile(r"(gas|petrol|energy|trade|industr|mining|غاز|بترول|طاقة|تجارة|صناع|تعدين)", re.I)
    cand = sorted({s[1] for s in subs if pat.search(s[2] + " " + (s[3] or ""))})
    out("  candidate ids:", cand)
    for sid in cand[:30]:
        rr = get(api + f"Subject/SubSubjectWithIndicator/{sid}", quiet=True)
        if rr is None or rr.status_code != 200:
            continue
        try:
            d = (rr.json().get("data") or {})
        except Exception:  # noqa: BLE001
            continue
        pubs = d.get("publicationWithIndicators") or []
        out(f"  subSubject {sid}: {len(pubs)} publications")
        for pb in pubs:
            if not isinstance(pb, dict):
                continue
            ptr = pb.get("publicationTranslations") or []
            pen = next((t.get("name") or t.get("title") for t in ptr if isinstance(t, dict) and t.get("locale") == "en"), "")
            out(f"    PUB {pb.get('id')} {str(pb.get('name') or pb.get('title') or '')[:60]} | {str(pen)[:60]} | "
                f"{ {k: pb.get(k) for k in ('periodicity', 'frequency', 'periodicityId') if k in pb} }")
            for k, v in pb.items():
                if isinstance(v, list):
                    for it in v:
                        if not isinstance(it, dict):
                            continue
                        nm = str(it.get("name") or it.get("title") or "")
                        tr = it.get("indicatorTranslations") or []
                        en = next((t.get("name") or t.get("title") for t in tr if isinstance(t, dict) and t.get("locale") == "en"), "")
                        if re.search(r"(gas|غاز|LNG|مسال|بترول|petrol)", nm + " " + str(en), re.I):
                            out(f"       ind {it.get('id')} {nm[:70]} | {str(en)[:70]}")


def run1():
    jodi()
    comtrade_probe()
    cbe()
    sites()
    capmas()
    wayback()


if __name__ == "__main__":
    {"1": run1}[ROUND]()
