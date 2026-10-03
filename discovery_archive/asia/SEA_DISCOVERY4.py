"""
South & Southeast Asia discovery, round 4 (after SEA_DISCOVERY3: EPPO T5.2-2M = monthly generation by fuel,
whole system; IEMOP admin-ajax returns base64 server paths such as
/var/www/html/wp-content/uploads/downloads/data/RTDREG/RTDREG_20261002.csv, about 90 days listed;
DEPS links are relative, /wp-content/uploads/2026/01/*.xlsx, and gave 404 when fetched cold).

  EPPO   T5.2-4M (generation by fuel, detail: does it split renewables?), T5.2-5M (peak), and the
         natural gas tables - NGV / natural-gas pages, T01_01_02 / 05 / 08 rows mentioning gas or LNG
  IEMOP  decode a few listed paths, download one RTD regional summary csv and one DIPCER zip
  DEPS   fetch the xlsx with a session that has loaded the eData page first
"""
import base64
import io
import re
import sys
import zipfile

import pandas as pd
import requests

from SEA_DISCOVERY_TH_VN_PH import H, out
from SEA_DISCOVERY2 import T, get, show_excel

U = "https://www.eppo.go.th/wp-content/uploads/2026/04/"


def eppo():
    for name in ("T05_02_04-1.xls", "T05_02_05-1.xls"):
        r = get(U + name)
        if r is not None and r.ok:
            df = pd.read_excel(io.BytesIO(r.content), header=None)
            out(df.head(12).to_string(max_cols=20, max_colwidth=16)[:4000])
            out(df.iloc[-40:-20].to_string(max_cols=20, max_colwidth=16)[:4000])
    for page in ("ngv", "natural-gas", "naturalgas", "gas", "coal", "petroleum-price", "energy-price", "lng"):
        r = get(f"https://www.eppo.go.th/epposite/info/stat/{page}")
        if r is not None and r.ok:
            xls = sorted(set(re.findall(r'href="([^"]+\.xlsx?)"', r.text)))
            out(f"  {page}: {len(xls)} xls: " + ", ".join(x.rsplit('/', 1)[-1] for x in xls[:80]))
    for name in ("T01_01_02-1.xls", "T01_01_05-1.xls", "T01_01_08-1.xls"):
        r = get(U + name)
        if r is not None and r.ok:
            df = pd.read_excel(io.BytesIO(r.content), header=None)
            hdr = df.head(10).fillna("").astype(str).agg(" | ".join, axis=1)
            out("  header rows:\n    " + "\n    ".join(hdr.tolist()))
            out(df.iloc[-30:-15].to_string(max_cols=20, max_colwidth=14)[:3000])


def iemop():
    for post, kind in ((5760, "RTDREG csv"), (5754, "DIPCER zip")):
        x = requests.post("https://www.iemop.ph/wp-admin/admin-ajax.php",
                          data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "", "page": 1,
                                "post_id": post}, headers=H, timeout=T)
        j = x.json()
        paths = [base64.b64decode(s).decode() for s in j.get("source", [])[:3]]
        out(f"  {kind}: count {j.get('count')} keys {list(j.keys())}; paths {paths}")
        out(f"  other fields: { {k: str(v)[:300] for k, v in j.items() if k != 'source'} }")
        rel = paths[0].split("/html", 1)[-1]
        for u in (f"https://www.iemop.ph{rel}",
                  f"https://www.iemop.ph/market-data/download/?file={j['source'][0]}",
                  f"https://www.iemop.ph/?md_file={j['source'][0]}"):
            r = get(u)
            if r is not None and r.ok and "html" not in (r.headers.get("content-type") or ""):
                if u.endswith(".zip") or r.content[:2] == b"PK":
                    z = zipfile.ZipFile(io.BytesIO(r.content))
                    out(f"  zip members {z.namelist()[:5]}")
                    out(z.read(z.namelist()[0]).decode(errors="replace")[:2500])
                else:
                    out(r.text[:2500])
                break


def deps():
    s = requests.Session()
    s.headers.update(H)
    p = s.get("https://www.deps.gov.bn/?p=1665", timeout=T)
    out(f"page {p.status_code} cookies {s.cookies.get_dict()}")
    for f in ("Gas-Production-LNG-Export-and-Weighted-Average-Price.xlsx", "Domestic-Consumption-of-Oil-and-Gas.xlsx"):
        for host in ("https://www.deps.gov.bn", "https://deps.mofe.gov.bn"):
            r = s.get(f"{host}/wp-content/uploads/2026/01/{f}", headers={"Referer": "https://www.deps.gov.bn/?p=1665"},
                      timeout=T)
            out(f"  {host} {f} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
            if r.ok and r.content[:2] == b"PK":
                show_excel(r.content, f, rows=30)
                break


if __name__ == "__main__":
    for f in sys.argv[1:] or ["eppo", "iemop", "deps"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
