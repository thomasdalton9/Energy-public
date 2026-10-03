"""
South & Southeast Asia discovery, round 2 (after SEA_DISCOVERY_TH_VN_PH / _MY_ID_BN, which found
from a GitHub Actions runner: EPPO's statistics page lists T05_xx xls tables under
wp-content/uploads/<yyyy>/<mm>/; RID app.rid.go.th/reservoir/api/dam/public returns today's 35
large dams; GSO page methods return 10-minute generation by fuel back to 2020; Single Buyer SMP API
answers; DEPS Brunei links gas/LNG xlsx files; IEMOP market-data pages load their file lists by
JavaScript; PAGASA's dam table is not in the static page; NSMO / EVN reservoir hosts time out).

This round opens the actual files:
  EPPO   T05_01_01 (capacity), T05_02_01 / -1 / -2 (generation by fuel: current year monthly,
         historical monthly, yearly?), T05_04_01 (peak), plus the NGV/natural-gas statistics page
         and its gas tables (supply, use by sector, LNG imports)
  RID    does the dam API take a date (history for the water-year chart)?
  IEMOP  the market-downloads plugin JS, to find its AJAX listing endpoint
  PAGASA dam information section / any JSON behind it
  DEPS   Brunei gas production / LNG export / price workbook layout
  Single Buyer  SMP with date parameters (actuals)
  Vietnam NSMO / EVN reservoirs with a longer timeout and http://
"""
import io
import re
import sys
from datetime import date, timedelta

import pandas as pd
import requests

from SEA_DISCOVERY_TH_VN_PH import H, out, probe

T = (20, 120)
EPPO = "https://www.eppo.go.th/wp-content/uploads/2026/04/"


def show_excel(content, label, rows=40, sheets=4):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        out(f"  {label}: not excel ({e!r}); head {content[:120]!r}")
        return
    out(f"  {label}: sheets {xl.sheet_names}")
    for s in xl.sheet_names[:sheets]:
        df = pd.read_excel(xl, s, header=None)
        out(f"  --- sheet {s!r} shape {df.shape}")
        out(df.head(rows).to_string(max_cols=16, max_colwidth=24)[:6000])
        out("  ... tail:")
        out(df.tail(6).to_string(max_cols=16, max_colwidth=24)[:2000])


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"\nGET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} "
            f"last-modified={r.headers.get('last-modified')}")
        return r
    except Exception as e:
        out(f"\nGET {url} ERROR {e!r}")
        return None


def eppo():
    for f in ("T05_01_01-2.xls", "T05_02_01.xls", "T05_02_01-1.xls", "T05_02_01-2.xls", "T05_04_01.xls",
              "T05_04_01-1.xls"):
        r = get(EPPO + f)
        if r is not None and r.ok:
            show_excel(r.content, f, rows=30, sheets=2)
    for u in ("https://www.eppo.go.th/epposite/info/stat/ngv", "https://www.eppo.go.th/epposite/info/stat/petroleum",
              "https://www.eppo.go.th/epposite/info/stat/summary", "https://www.eppo.go.th/epposite/info/stat/price"):
        r = probe("EPPO stat page", u, show=1200, links=150)


def rid():
    d = (date.today() - timedelta(days=400)).isoformat()
    for u in (f"https://app.rid.go.th/reservoir/api/dam/public?date={d}",
              f"https://app.rid.go.th/reservoir/api/dam/public/{d}",
              "https://app.rid.go.th/reservoir/api/document/dam",
              "https://app.rid.go.th/reservoir/api/dam/public?date=2020-01-15"):
        r = get(u)
        if r is not None:
            out(f"  {r.text[:400]}")
    r = get("https://app.rid.go.th/reservoir/")
    if r is not None:
        for js in re.findall(r'src="([^"]+\.js[^"]*)"', r.text)[:8]:
            j = get(js if js.startswith("http") else "https://app.rid.go.th" + js)
            if j is not None:
                out("  api paths: " + ", ".join(sorted(set(re.findall(r'["\'](/?[\w/]*api/[\w/{}$.-]+)', j.text)))[:40]))


def iemop():
    r = get("https://www.iemop.ph/wp-content/plugins/iemop-market-downloads/public/js/market-downloads-public.js?ver=5.5.3")
    if r is not None:
        out(r.text[:6000])
    r = get("https://www.iemop.ph/market-data/rtd-regional-summaries/")
    if r is not None:
        for m in re.finditer(r"(ajax[^\n]{0,200}|data-[\w-]+=\"[^\"]{0,120}\")", r.text):
            out("  " + m.group(1))


def pagasa():
    r = get("https://www.pagasa.dost.gov.ph/flood")
    if r is not None:
        i = r.text.lower().find("dam")
        for m in re.finditer(r"(?is)<table.*?</table>", r.text):
            t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", m.group(0)))
            if "angat" in t.lower() or "dam" in t.lower():
                out("  TABLE: " + t[:4000])
        out("  dam-ish urls: " + ", ".join(sorted(set(re.findall(r'["\']([^"\']*dam[^"\']*)["\']', r.text, re.I)))[:40]))


def deps():
    for f in ("Gas-Production-LNG-Export-and-Weighted-Average-Price.xlsx", "Domestic-Consumption-of-Oil-and-Gas.xlsx"):
        r = get("https://www.deps.gov.bn/wp-content/uploads/2026/01/" + f)
        if r is not None and r.ok:
            show_excel(r.content, f, rows=25)


def single_buyer():
    d1 = date.today() - timedelta(days=3)
    for q in (f"?date={d1.isoformat()}", f"?startDate={d1.isoformat()}&endDate={d1.isoformat()}",
              f"?from={d1.isoformat()}&to={d1.isoformat()}", "?date=2024-01-15"):
        r = get("https://www.singlebuyer.com.my/api/v1/smp/actual-forecast" + q)
        if r is not None:
            try:
                j = r.json()
                a = j.get("meta", {}).get("data", {}).get("actual") or j.get("data", {}).get("actual")
                out(f"  meta {str(j.get('meta'))[:200]}; actual n={len(a or [])} first {(a or [])[:2]}")
            except Exception:
                out(f"  {r.text[:300]}")
    r = get("https://www.singlebuyer.com.my/")
    if r is not None:
        out("  api paths: " + ", ".join(sorted(set(re.findall(r'(/api/v1/[\w/-]+)', r.text)))[:40]))


def vietnam():
    for u in ("http://www.nsmo.vn/", "https://nsmo.vn/", "http://hochuathuydien.evn.com.vn/"):
        get(u)


if __name__ == "__main__":
    for f in sys.argv[1:] or ["eppo", "rid", "iemop", "pagasa", "deps", "single_buyer", "vietnam"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
