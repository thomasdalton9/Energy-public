"""
Capacity + price discovery for South & Southeast Asia, round 1 (GitHub runner). Probes:
  Singapore  NEMS DataDownload value=1 (USEP + demand, half-hourly CSV) and DataDownloadByYear value=1
  India      CEA installed-capacity page links, latest IC xlsx sheets; MNRE physical progress; NPP installcap
  Thailand   EPPO electricity page: T05_01_* capacity tables
  Bangladesh BPDB power-generation-unit and home page capacity tables; daily-generation report capacity
  Philippines DOE power statistics / list of existing power plants pages
  Vietnam    NSMO / EVN reachability for market price and capacity
  Sri Lanka  PUCSL metadata/power-plants capacity sum
"""
import io
import re
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 120)
pd.set_option("display.width", 250)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} LM={r.headers.get('last-modified')}")
        return r
    except Exception as e:
        out(f"GET {u} !! {type(e).__name__}: {str(e)[:200]}")
        return None


def links(r, pat):
    if r is None:
        return []
    ls = sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text)))
    return [l for l in ls if re.search(pat, l, re.I)]


def show_xl(content, rows=40, sheets=None):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        out(f"  not excel: {e}")
        return
    out(f"  sheets: {xl.sheet_names}")
    for s in (sheets or xl.sheet_names[:3]):
        if s not in xl.sheet_names:
            continue
        df = xl.parse(s, header=None)
        out(f"  --- {s} {df.shape}")
        out(df.head(rows).to_string(max_cols=16, max_colwidth=22))


def singapore():
    out("\n######## SINGAPORE")
    N = "https://www.nems.emcsg.com/api/sitecore/DataSync"
    r = get(f"{N}/DataDownload", params={"value": "1", "fromDate": "2026-09-01", "toDate": "2026-09-02", "tpcValue": "1"})
    if r is not None:
        out(r.text[:1500])
    r = get(f"{N}/DataDownloadByYear", params={"value": "1", "year": "2021", "tpcValue": "1"})
    if r is not None and r.content[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(r.content))
        out(z.namelist()[:20])
        t = z.read(z.namelist()[0]).decode("utf-8-sig", "replace")
        out(t[:800])
    r = get(f"{N}/DataDownloadByYear", params={"value": "1", "year": "2026", "tpcValue": "1"})
    if r is not None and r.content[:2] == b"PK":
        out(zipfile.ZipFile(io.BytesIO(r.content)).namelist())


def india():
    out("\n######## INDIA")
    r = get("https://cea.nic.in/installed-capacity-report/?lang=en")
    ls = links(r, r"installed|IC_|capacity")
    out(f"  {len(ls)} links"); [out("   ", l) for l in ls[:80]]
    xl = [l for l in ls if l.lower().endswith((".xlsx", ".xls"))]
    for l in xl[:2]:
        rr = get(l)
        if rr is not None and rr.status_code == 200:
            show_xl(rr.content, rows=60)
    # older months: try direct names
    for d in (date(2026, 8, 31), date(2025, 3, 31), date(2023, 3, 31), date(2021, 3, 31)):
        get(f"https://cea.nic.in/wp-content/uploads/installed/{d.year}/{d.month:02d}/IC_allocation_as_on_{d:%d.%m.%Y}.xlsx")
    for u in ("https://cea.nic.in/wp-admin/admin-ajax.php?action=installed_capacity",
              "https://cea.nic.in/api/installed_capacity.php",
              "https://cea.nic.in/dashboard/?lang=en"):
        rr = get(u)
        if rr is not None:
            out(rr.text[:400])
    r = get("https://mnre.gov.in/en/physical-progress/")
    if r is not None and r.status_code == 200:
        try:
            for t in pd.read_html(io.StringIO(r.text))[:3]:
                out(t.head(40).to_string())
        except Exception as e:
            out(f"  no tables: {e}")
            out(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:3000])
    r = get("https://npp.gov.in/publishedReports")
    for l in links(r, r"install|capac"):
        out("   ", l)


def thailand():
    out("\n######## THAILAND")
    r = get("https://www.eppo.go.th/epposite/info/stat/electricity")
    ls = links(r, r"T05_01")
    for l in ls:
        out("   ", l)
    for l in ls:
        rr = get(l if l.startswith("http") else "https://www.eppo.go.th" + l)
        if rr is not None and rr.status_code == 200:
            show_xl(rr.content, rows=45)


def bangladesh():
    out("\n######## BANGLADESH")
    for u in ("https://misc.bpdb.gov.bd/power-generation-unit", "https://bpdb.gov.bd/",
              "https://misc.bpdb.gov.bd/", "https://bpdb.gov.bd/site/page/installed-capacity"):
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        try:
            ts = pd.read_html(io.StringIO(r.text))
            out(f"  {len(ts)} tables")
            for t in ts[:4]:
                out(t.head(25).to_string(max_cols=14, max_colwidth=24)); out(f"   ... {t.shape}")
        except Exception as e:
            out(f"  no tables: {e}")
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        for m in re.finditer(r"(?i)installed|capacity", txt):
            out("   ctx:", txt[max(0, m.start() - 150): m.start() + 250]); break
    d = date.today() - timedelta(days=3)
    r = get(f"https://misc.bpdb.gov.bd/daily-generation?date={d:%d-%m-%Y}")
    if r is not None and r.status_code == 200:
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        for m in list(re.finditer(r"(?i)installed|derated|capacity", txt))[:6]:
            out("   ctx:", txt[max(0, m.start() - 200): m.start() + 400])


def philippines():
    out("\n######## PHILIPPINES")
    for u in ("https://doe.gov.ph/energy-information-resources?q=electric-power/power-statistics",
              "https://doe.gov.ph/electric-power/list-existing-power-plants",
              "https://legacy.doe.gov.ph/electric-power-statistics",
              "https://legacy.doe.gov.ph/list-existing-power-plants",
              "https://legacy.doe.gov.ph/energy-statistics",
              "https://www.doe.gov.ph/energy-statistics"):
        r = get(u)
        for l in links(r, r"\.xlsx?$|\.pdf$|statistic|existing|power-plant")[:40]:
            out("   ", l)


def vietnam():
    out("\n######## VIETNAM")
    for u in ("https://www.nsmo.vn/", "http://www.nsmo.vn/", "https://vwem.nsmo.vn/", "https://www.evn.com.vn/c3/evn-va-khach-hang/Gia-dien-9-79.aspx",
              "https://www.evn.com.vn/vi-VN/news-l/Bao-cao-thuong-nien-60-17"):
        r = get(u)
        if r is not None:
            out("  ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:300])


def srilanka():
    out("\n######## SRI LANKA")
    r = get("https://gendata.pucsl.gov.lk/api/metadata/power-plants",
            headers={"Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"})
    if r is not None and r.status_code == 200:
        d = r.json().get("data", [])
        out(f"  {len(d)} plants; keys {list(d[0].keys()) if d else None}")
        out(d[0] if d else None)


if __name__ == "__main__":
    for f in (singapore, india, thailand, bangladesh, philippines, vietnam, srilanka):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__}: {e!r}")
