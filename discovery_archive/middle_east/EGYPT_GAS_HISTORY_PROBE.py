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


# ------------------------------------------------------------------ round 2
GASPAT = re.compile(r"(natural gas|\bgas\b|petroleum|LNG|غاز|بترول|بوتاجاز)", re.I)


def xlsx_scan(url, maxrows=25):
    import pandas as pd
    out(f"\n=== XLSX {url}")
    r = get(url)
    if r is None or r.status_code != 200:
        out("  failed", None if r is None else r.status_code)
        return None
    out(f"  {len(r.content)}B {r.headers.get('content-type','')[:40]}")
    try:
        xl = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
    except Exception as e:  # noqa: BLE001
        out("  read error", type(e).__name__, str(e)[:150])
        return None
    out("  sheets:", list(xl.keys())[:80])
    for name, df in xl.items():
        txt = df.astype(str)
        hits = [i for i in range(len(df)) if GASPAT.search(" ".join(txt.iloc[i].tolist()))]
        title = " ".join(x for x in txt.iloc[:4].values.ravel() if x != "nan")[:160]
        if hits or GASPAT.search(name + title):
            out(f"  -- sheet {name!r} shape {df.shape} title: {title}")
            for i in range(min(6, len(df))):
                out("     hdr", [x[:18] for x in txt.iloc[i].tolist() if x != "nan"][:16])
            for i in hits[:maxrows]:
                out("     row", i, [x[:16] for x in txt.iloc[i].tolist() if x != "nan"][:18])
    return xl


def cbe2():
    out("\n################ CBE monthly statistical bulletin xlsx")
    base = "https://www.cbe.org.eg/-/media/project/cbe/listing/monthly-statistical-bulletin/"
    xlsx_scan(base + "real/real-sector-352.xlsx")
    xlsx_scan(base + "external/external-sector-352.xlsx", maxrows=12)
    out("\n--- CBE time series lists")
    for cat in ["F0324992E95741438C789A669E5194F4", "232131B16F15454BB1E1933B2BFEB041", "DEF6421CA1354B128A1113D7A5BBFC66"]:
        page(f"https://www.cbe.org.eg/en/economic-research/time-series/downloadlist?category={cat}", max_links=40,
             pat=re.compile(r"(xls|pdf|csv|media)", re.I))


def wb_get(ts, url):
    w = f"http://web.archive.org/web/{ts}id_/{url}"
    for att in range(3):
        try:
            r = requests.get(w, headers=H, timeout=(20, 120))
            if r.status_code in (429, 503):
                time.sleep(15 * (att + 1))
                continue
            return r
        except Exception as e:  # noqa: BLE001
            out("  WB ERROR", type(e).__name__, str(e)[:100])
            time.sleep(8)
    return None


def text_of(r):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S | re.I)
    return " ".join(re.sub(r"<[^>]+>", " ", t).split())


def pdf_text(content, maxpages=400):
    import fitz
    doc = fitz.open(stream=content, filetype="pdf")
    return doc, [doc[i].get_text() for i in range(min(len(doc), maxpages))]


def pdf_scan(name, content, pat=GASPAT, ctx=220, maxhits=25):
    try:
        doc, pages = pdf_text(content)
    except Exception as e:  # noqa: BLE001
        out("  pdf error", type(e).__name__, str(e)[:120])
        return
    nchars = sum(len(p) for p in pages)
    out(f"  PDF {name}: {len(doc)} pages, {nchars} text chars")
    n = 0
    for i, t in enumerate(pages):
        for m in re.finditer(r"(bcf|billion cubic|BCM|mmscf|million cubic|مليار متر|مليون متر|قدم مكعب|شحن|cargo)", t, re.I):
            s = " ".join(t[max(0, m.start() - ctx):m.end() + ctx].split())
            out(f"   p{i+1}: ...{s}...")
            n += 1
            if n >= maxhits:
                return


def egpc_mop():
    out("\n################ EGPC (http) live + Wayback")
    for u in ["http://www.egpc.com.eg/", "http://www.egpc.com.eg/Production.aspx", "http://www.egpc.com.eg/Production_2.aspx",
              "http://www.egpc.com.eg/Gas_Production.aspx", "http://www.egpc.com.eg/Annual_Reports.aspx",
              "http://www.egpc.com.eg/AnnualReportYears.aspx", "http://www.egpc.com.eg/E/Default.aspx",
              "http://www.egpc.com.eg/DistMainProducts.aspx"]:
        r = page(u, max_links=40)
        if r is not None and r.status_code == 200 and "html" in r.headers.get("content-type", ""):
            t = text_of(r)
            for m in list(re.finditer(r"(bcf|cubic|مكعب|gas|غاز)", t, re.I))[:6]:
                out("    txt:", t[max(0, m.start() - 150):m.end() + 200])
    for ts, u in [("20260309121026", "http://www.egpc.com.eg/Production.aspx"),
                  ("20260309120813", "http://www.egpc.com.eg/Production_2.aspx"),
                  ("20180731141455", "http://www.egpc.com.eg:80/Gas_Production.aspx"),
                  ("20180602225604", "http://www.egpc.com.eg:80/Annual_Reports.aspx"),
                  ("20210419000604", "http://egpc.com.eg/Annual_Reports_14.aspx")]:
        out(f"\n=== WB {ts} {u}")
        r = wb_get(ts, u)
        if r is None:
            continue
        out("  ", r.status_code, len(r.content))
        t = text_of(r)
        out("   text:", t[:1500])
        links = sorted(set(re.findall(r'href=["\']([^"\']+\.(?:pdf|xlsx?))["\']', r.text, re.I)))
        out("   file links:", links[:40])
    for ts, u in [("20250503172442", "http://www.egpc.com.eg/E/PDF/production/Oil%20%20Gas5-2-2017.pdf"),
                  ("20191213184725", "http://www.egpc.com.eg/PDF/production/Production%20Activity.pdf")]:
        r = wb_get(ts, u)
        out(f"\n=== WB PDF {u} -> {None if r is None else r.status_code} {0 if r is None else len(r.content)}")
        if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
            pdf_scan(u, r.content)

    out("\n################ Ministry of Petroleum info-indicators")
    for u in ["https://www.petroleum.gov.eg/ar-eg/gas-and-petrol/Pages/info-indicators.aspx",
              "https://www.petroleum.gov.eg/en/gas-and-petrol/Pages/info-indicators.aspx",
              "https://www.petroleum.gov.eg/ar-eg/gasand-petrol/Pages/info-indicators.aspx",
              "https://www.petroleum.gov.eg/ar-eg/Pages/default.aspx",
              "https://www.petroleum.gov.eg/ar-eg/media-center/news/Pages/default.aspx"]:
        r = page(u, max_links=30)
        if r is not None and r.status_code == 200:
            t = text_of(r)
            for m in list(re.finditer(r"(مليار قدم|مليون قدم|bcf|مكعب|cubic)", t, re.I))[:10]:
                out("    txt:", t[max(0, m.start() - 200):m.end() + 200])
    for ts, u in [("20200422020132", "https://www.petroleum.gov.eg/ar-eg/gas-and-petrol/Pages/info-indicators.aspx"),
                  ("20240831190814", "https://www.petroleum.gov.eg/ar-eg/gasand-petrol/Pages/info-indicators.aspx")]:
        r = wb_get(ts, u)
        out(f"\n=== WB {ts} {u} -> {None if r is None else r.status_code}")
        if r is not None and r.status_code == 200:
            t = text_of(r)
            i = max(t.find("مؤشر"), 0)
            out("   text:", t[i:i + 2500])
    rows = cdx("www.petroleum.gov.eg/*", limit=400, filt="urlkey:.*(indicator|statistic|ehsa|production|gas-and-petrol/pages).*")
    for row in rows[:60]:
        out("   ", row[1], row[3][:22], row[2][:170])


def egas2():
    out("\n################ EGAS annual reports")
    for u in ["https://www.egas.com.eg/annual-report", "https://www.egas.com.eg/ar/annual-report",
              "https://www.egas.com.eg/en/annual-report", "https://www.egas.com.eg/activities/production-and-field-development",
              "https://www.egas.com.eg/w/production-and-development", "https://www.egas.com.eg/ar/w/production-and-development"]:
        r = page(u, max_links=40, pat=re.compile(r"(annual|report|documents|pdf|تقرير|سنوي)", re.I))
        if r is not None and r.status_code == 200:
            t = text_of(r)
            for m in list(re.finditer(r"(bcf|cubic|مكعب|billion)", t, re.I))[:8]:
                out("    txt:", t[max(0, m.start() - 200):m.end() + 200])
    for ts, u in [("20191022142621", "https://www.egas.com.eg/annual-reports/2018"),
                  ("20191115075205", "https://www.egas.com.eg/annual-reports/2019"),
                  ("20220628211320", "https://www.egas.com.eg/annual-report"),
                  ("20260912065513", "https://egas.com.eg/ar/%D8%A7%D9%84%D8%B1%D8%A6%D9%8A%D8%B3%D9%8A%D8%A9/-/asset_publisher/fezn/content/annual-report-2022?p_r_p_assetEntryId=96751"),
                  ("20220628211319", "https://www.egas.com.eg/activities/production")]:
        out(f"\n=== WB {ts} {u}")
        r = wb_get(ts, u)
        if r is None:
            continue
        out("  ", r.status_code, len(r.content))
        t = text_of(r)
        for m in list(re.finditer(r"(bcf|cubic|مكعب|billion|BCM)", t, re.I))[:10]:
            out("    txt:", t[max(0, m.start() - 200):m.end() + 200])
        links = sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+(?:\.pdf|/documents/[^"\']+))["\']', r.text, re.I)))
        out("   file links:", links[:30])
    for ts, u in [("20160610162620", "http://egas.com.eg/docs/pdf/Egas_EN_2015_Final%20@3-9-2015.pdf"),
                  ("20160902095431", "http://www.egas.com.eg/docs/pdf/AR_mid.pdf"),
                  ("20150202113402", "http://www.egas.com.eg/docs/pdf/EN_mid.pdf")]:
        r = wb_get(ts, u)
        out(f"\n=== WB PDF {u} -> {None if r is None else r.status_code} {0 if r is None else len(r.content)}")
        if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
            pdf_scan(u, r.content)
    rows = cdx("egas.com.eg/documents/*", limit=300)
    for row in rows[:80]:
        out("   ", row[1], row[3][:22], row[2][:200])


def capmas2():
    out("\n################ CAPMAS indicator data (energy statistics pub 77, trade pubs)")
    api = "https://www.capmas.gov.eg:8080/api/Indicator/IndicatorFilter?IndicatorId={}&SubSubjectId={}"
    for iid, ss in [(1665706, 30), (1664307, 30), (1664313, 30), (1664212, 30), (1664348, 30), (1664275, 30),
                    (962760, 30), (962821, 30), (963034, 30), (962673, 30), (908278, 10), (908592, 10)]:
        r = get(api.format(iid, ss), quiet=True)
        if r is None:
            out(f"  {iid}: no response")
            continue
        try:
            rows = r.json().get("data") or []
        except Exception:  # noqa: BLE001
            out(f"  {iid}: {r.status_code} not json {r.text[:120]}")
            continue
        out(f"\n  indicator {iid}: {len(rows)} filters")
        for f in rows[:40]:
            en = next((t.get("name") for t in f.get("filterTranslations", []) if t.get("locale") == "en"), "")
            pts = sorted((d.get("year") or 0, d.get("quarter") or 0, d.get("month") or 0, d.get("value")) for d in f.get("data", []))
            out(f"    {str(f.get('name'))[:45]} | {str(en)[:45]} | n={len(pts)} first={pts[:1]} last={pts[-2:]}")
    out("\n  -- all indicators of trade pubs 22 / 100 and production pub 18")
    for ss in (10, 12):
        rr = get(f"https://www.capmas.gov.eg:8080/api/Subject/SubSubjectWithIndicator/{ss}", quiet=True)
        try:
            pubs = (rr.json().get("data") or {}).get("publicationWithIndicators") or []
        except Exception:  # noqa: BLE001
            continue
        for pb in pubs:
            if pb.get("id") not in (22, 100, 18):
                continue
            for k, v in pb.items():
                if isinstance(v, list):
                    for it in v[:120]:
                        if isinstance(it, dict) and (it.get("name") or it.get("title")):
                            tr = it.get("indicatorTranslations") or []
                            en = next((t.get("name") for t in tr if isinstance(t, dict) and t.get("locale") == "en"), "")
                            out(f"    pub{pb.get('id')} ind {it.get('id')} {str(it.get('name'))[:60]} | {str(en)[:70]}")


def comtrade2():
    out("\n################ Comtrade partner detail (includeDesc)")
    for y in [2016, 2024, 2025]:
        rows = comtrade("A", reporterCode=818, period=y, cmdCode="271111,271121", flowCode="M,X", includeDesc="true")
        for x in sorted(rows or [], key=lambda x: -(x.get("netWgt") or 0))[:10]:
            out(f"   {y} {x.get('flowCode')} {x.get('cmdCode')} {x.get('partnerDesc')} t={(x.get('netWgt') or 0)/1000:,.0f} "
                f"USD_m={(x.get('primaryValue') or 0)/1e6:,.1f} qtyUnit={x.get('qtyUnitAbbr')} qty={x.get('qty')}")
    for p in ["202501", "202604", "202605", "202606", "202607", "202608"]:
        rows = comtrade("M", reporterCode=818, period=p, cmdCode="271111,271121", flowCode="M,X", includeDesc="true")
        out(f"  {p}: {None if rows is None else len(rows)} rows")
        for x in sorted(rows or [], key=lambda x: -(x.get("netWgt") or 0))[:6]:
            out(f"   {p} {x.get('flowCode')} {x.get('cmdCode')} {x.get('partnerDesc')} t={(x.get('netWgt') or 0)/1000:,.0f}")
    out("  -- monthly availability scan 2010-2026 (271111 only, one call per month, M+X world)")
    have = []
    for y in range(2010, 2027):
        for m in range(1, 13):
            p = f"{y}{m:02d}"
            if p > "202608":
                break
            if m not in (1, 7):
                continue
            rows = comtrade("M", reporterCode=818, period=p, cmdCode="271111", flowCode="M,X", partnerCode=0)
            have.append((p, None if rows is None else len(rows),
                         sum((x.get("netWgt") or 0) for x in rows or [] if x.get("flowCode") == "M") / 1000,
                         sum((x.get("netWgt") or 0) for x in rows or [] if x.get("flowCode") == "X") / 1000))
    for h in have:
        out("   ", h)


def run2():
    cbe2()
    capmas2()
    egpc_mop()
    egas2()
    comtrade2()


if __name__ == "__main__":
    {"1": run1, "2": run2}[ROUND]()
