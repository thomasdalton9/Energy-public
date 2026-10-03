"""
South & Southeast Asia discovery, round 3 (after SEA_DISCOVERY2: EPPO T5.2-1 splits fuel only for
EGAT's own plants (IPP/SPP/VSPP by producer), so the by-fuel table is elsewhere; IEMOP lists files via
admin-ajax action display_filtered_market_data_files (post_id per page, min_date about 5 months back);
DEPS xlsx links returned 404 when fetched directly).

  EPPO   title rows of every T01 / T02 / T05 monthly table (-1 files) to find generation by fuel,
         natural gas supply / use by sector / LNG imports, then dump the chosen ones
  IEMOP  admin-ajax listing for a few market-data pages
  DEPS   eData page links again, fetched with the page as Referer
"""
import re
import sys

import requests

from SEA_DISCOVERY_TH_VN_PH import H, out
from SEA_DISCOVERY2 import T, get, show_excel

import pandas as pd
import io

EPPO = "https://www.eppo.go.th/epposite/info/stat/"


def eppo():
    files = {}
    for page in ("electricity", "petroleum", "summary"):
        r = get(EPPO + page)
        if r is None:
            continue
        for u in sorted(set(re.findall(r'href="([^"]+/T0\d_\d\d_\d\d(?:-\d)?\.xls)"', r.text))):
            files[u.rsplit("/", 1)[-1]] = u
    out(f"{len(files)} tables")
    for name in sorted(files):
        if not name.endswith("-1.xls"):
            continue
        try:
            x = requests.get(files[name], headers=H, timeout=T)
            df = pd.read_excel(io.BytesIO(x.content), header=None, nrows=8)
            title = " | ".join(str(v) for v in df.iloc[:4, :3].values.ravel() if str(v) != "nan")
            out(f"  {name}: {title[:160]}")
        except Exception as e:
            out(f"  {name}: {e!r}")
    for name in ("T05_02_02-1.xls", "T05_02_03-1.xls", "T05_03_01-1.xls"):
        if name in files:
            r = get(files[name])
            if r is not None and r.ok:
                show_excel(r.content, name, rows=12, sheets=1)
                df = pd.read_excel(io.BytesIO(r.content), header=None)
                out(df.iloc[-60:-30].to_string(max_cols=16, max_colwidth=14)[:5000])


def iemop():
    for slug in ("rtd-regional-summaries", "dipc-energy-results-raw", "rtd-market-prices-lmp", "rtd-zonal-prices",
                 "reserve-market-prices"):
        r = get(f"https://www.iemop.ph/market-data/{slug}/")
        if r is None or not r.ok:
            continue
        m = re.search(r'"post_id":"(\d+)","min_date":"([^"]+)"', r.text)
        if not m:
            continue
        out(f"  {slug}: post_id {m.group(1)} min_date {m.group(2)}")
        x = requests.post("https://www.iemop.ph/wp-admin/admin-ajax.php",
                          data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                "page": 1, "post_id": m.group(1)}, headers=dict(H, Referer=r.url), timeout=T)
        out(f"  ajax -> {x.status_code} {len(x.content)} {x.text[:1500]}")


def deps():
    r = get("https://www.deps.gov.bn/?p=1665")
    if r is None:
        return
    links = sorted(set(re.findall(r'href="([^"]+\.xlsx?)"', r.text)))
    out("  xlsx: " + "\n    ".join(links[:80]))
    for u in [l for l in links if re.search("gas|lng|consumption", l, re.I)][:3]:
        x = requests.get(u, headers=dict(H, Referer="https://www.deps.gov.bn/?p=1665"), timeout=T)
        out(f"  {u} -> {x.status_code} {x.headers.get('content-type')} {len(x.content)}")
        if x.ok and "html" not in (x.headers.get("content-type") or ""):
            show_excel(x.content, u.rsplit("/", 1)[-1], rows=25)


if __name__ == "__main__":
    for f in sys.argv[1:] or ["eppo", "iemop", "deps"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
