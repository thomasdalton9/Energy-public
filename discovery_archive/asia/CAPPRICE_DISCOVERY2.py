"""
Capacity discovery round 2 (after CAPPRICE_DISCOVERY1: NEMS USEP CSV value=1 works (yearly zips + 31-day
ranges); CEA IC_allocation xlsx 'Summary' = all-India MW by fuel (latest month only on the page; older names 404);
NPP publishedReports lists /public-reports/cea/monthly/installcap/YYYY/MON/capacity1-YYYY-MM.xls; EPPO has only
Table 5.1-1Y (annual MW by producer EGAT/IPP/SPP/Imported); BPDB power-generation-unit loads its table by ajax;
DOE pages point to /data-and-prices/energy-statistics; nsmo.vn times out).
  India      NPP capacity1 layout (latest + 2021-01 + 2023-03), how far back it goes
  Bangladesh BPDB power-generation-unit ajax URL + response; daily-generation for older dates
  Philippines DOE /data-and-prices/energy-statistics links
  Thailand   every T05_ link on the EPPO electricity page
  Sri Lanka  plant metadata
"""
import io
import re
from datetime import date

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 120)
pd.set_option("display.width", 250)
MON = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    hd = dict(H, **kw.pop("headers", {}))
    try:
        r = requests.get(u, headers=hd, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} LM={r.headers.get('last-modified')}")
        return r
    except Exception as e:
        out(f"GET {u} !! {type(e).__name__}: {str(e)[:200]}")
        return None


def text(h):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))


def india():
    out("\n######## INDIA NPP capacity1")
    for y, m, show in ((2026, 8, True), (2021, 1, True), (2023, 3, False), (2019, 1, False), (2017, 3, False),
                       (2024, 6, False), (2025, 11, False)):
        u = f"https://npp.gov.in/public-reports/cea/monthly/installcap/{y}/{MON[m-1]}/capacity1-{y}-{m:02d}.xls"
        r = get(u)
        if r is not None and r.status_code == 200 and show:
            try:
                df = pd.read_excel(io.BytesIO(r.content), header=None)
                out(df.shape)
                out(df.head(70).to_string(max_cols=20, max_colwidth=18))
            except Exception as e:
                out(f"  parse: {e}; head {r.content[:300]!r}")


def bangladesh():
    out("\n######## BANGLADESH")
    r = get("https://misc.bpdb.gov.bd/power-generation-unit")
    if r is not None:
        i = r.text.find("capacity_type")
        out(r.text[i - 200:i + 1800])
        for m in re.findall(r"url\s*:\s*['\"]([^'\"]+)['\"]", r.text):
            out("  ajax url:", m)
            for ct in ("1", "2", "installed", "derated"):
                rr = get(m if m.startswith("http") else "https://misc.bpdb.gov.bd/" + m.lstrip("/"),
                         params={"capacity_type": ct}, headers={"X-Requested-With": "XMLHttpRequest"})
                if rr is not None:
                    out(text(rr.text)[:2500])
        for m in re.findall(r'<option[^>]*value="([^"]*)"[^>]*>(.*?)</option>', r.text):
            out("  option", m)
    for d in ("15-09-2026", "30-09-2026", "01-01-2026"):
        rr = get(f"https://misc.bpdb.gov.bd/daily-generation?date={d}")
        if rr is not None:
            t = text(rr.text)
            i = t.find("Daily Generation")
            out(t[i:i + 2500])


def philippines():
    out("\n######## PHILIPPINES")
    r = get("https://doe.gov.ph/data-and-prices/energy-statistics")
    if r is not None:
        ls = sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', r.text)))
        for l in ls:
            if re.search(r"stat|power|plant|xls|pdf|api|capacity", l, re.I):
                out("   ", l)
        out(text(r.text)[:3000])
        for m in sorted(set(re.findall(r'["\'](/api/[^"\']+|https?://[^"\']*api[^"\']*)["\']', r.text)))[:30]:
            out("  api:", m)


def thailand():
    out("\n######## THAILAND")
    r = get("https://www.eppo.go.th/energy-statistics/electricity-statistic/")
    if r is not None:
        for l in sorted(set(re.findall(r'href=["\']([^"\']+T0[0-9]_[0-9_\-]+\.[a-z]+)["\']', r.text))):
            out("   ", l)
        t = text(r.text)
        for m in re.finditer(r"(?i)capacity|กำลังผลิต", t):
            out("   ctx:", t[max(0, m.start() - 120): m.start() + 200])


def srilanka():
    out("\n######## SRI LANKA")
    r = get("https://gendata.pucsl.gov.lk/api/metadata/power-plants",
            headers={"Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"})
    if r is not None and r.status_code == 200:
        d = r.json().get("data", [])
        out(f"  {len(d)} plants")
        out(d[0] if d else None)


if __name__ == "__main__":
    for f in (india, bangladesh, philippines, thailand, srilanka):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__}: {e!r}")
