"""
South Asia discovery, round 3 (after SEA_DISCOVERY_SOUTH_ASIA2: NPP dgr1 = all-India/regional thermal,
nuclear, hydro daily MU; dgr2 = plant level with TYPE subtotal rows; CEA daily RE xlsx listing stops at
2025-11-18 (monthly RE PDFs since); CEA IC_allocation xlsx 'Summary' sheet = all-India MW by fuel;
PPAC current-FY xlsx for production / sectoral consumption / LNG imports; PGCB table cells hold no
'dd-mm-yyyy hh:mm:ss' string).

  PPAC    AjaxController getGasProduction / getGasConsumption for older financial years; the history
          (NG-H) sectoral file; NG-H LNG import 'Month wise' sheet layout
  NPP     full header rows of dgr2, dgr6 and dailyCoal1 (all columns) plus their all-India rows; how far
          back dgr1 / dailyCoal1 go (2021, 2019)
  CEA     older IC_allocation file names
  IGX     cumulative GIXI JSON
  PGCB    raw HTML of the first table rows
"""
import io
import os
import re
import sys
from datetime import date, timedelta

import pandas as pd
import requests
import urllib3

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "asia"))
from SEA_DISCOVERY_TH_VN_PH import H, out  # noqa: E402
from SEA_DISCOVERY2 import T, get, show_excel  # noqa: E402

urllib3.disable_warnings()
pd.set_option("display.width", 400)


def ppac():
    s = requests.Session()
    s.headers.update(H)
    s.get("https://ppac.gov.in/natural-gas/production", timeout=T)
    for path, page, fy in (("getGasProduction", 170, "2021-2022"), ("getGasProduction", 170, "2018-2019"),
                           ("getGasConsumption", 138, "2022-2023")):
        r = s.post(f"https://ppac.gov.in/AjaxController/{path}", data={"financialYear": fy, "reportBy": 4, "pageId": page},
                   headers={"X-Requested-With": "XMLHttpRequest", "Referer": "https://ppac.gov.in/natural-gas/production"},
                   timeout=T)
        try:
            res = r.json().get("result", {})
            out(f"  {path} {fy}: {len(res)} rows; titles {[v.get('title') for v in res.values()][:40]}")
            out(f"  sample {list(res.values())[-1]}")
        except Exception:
            out(f"  {path} {fy}: {r.status_code} {r.text[:300]}")
    for u in ("https://ppac.gov.in/uploads/page-images/1781859871_2_NG-H_Sectoral_Consumption.xlsx",
              "https://ppac.gov.in/uploads/pages/1781859871_2_NG-H_Sectoral_Consumption.xlsx"):
        r = get(u)
        if r is not None and r.ok and r.content[:2] == b"PK":
            show_excel(r.content, "NG-H sectoral", rows=25, sheets=2)
            break
    r = get("https://ppac.gov.in/uploads/page-images/1781164918_NG-H_LNG_Import.xlsx")
    if r is not None and r.ok:
        xl = pd.ExcelFile(io.BytesIO(r.content))
        for sh in xl.sheet_names[:2] + xl.sheet_names[-1:]:
            out(f"  --- {sh}\n" + pd.read_excel(xl, sh, header=None).head(25).to_string(max_cols=16, max_colwidth=20))
    for pg in ("production", "consumption", "sectoral-consumption", "import"):
        r = get(f"https://ppac.gov.in/natural-gas/{pg}")
        if r is not None:
            out(f"  {pg} xls links: " + ", ".join(sorted(set(re.findall(r'["\']([^"\']+\.xlsx?)["\']', r.text)))))


def npp():
    d = date.today() - timedelta(days=3)
    dd, iso = d.strftime("%d-%m-%Y"), d.isoformat()
    for name, url in (("dgr2", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr2-{iso}.xls"),
                      ("dgr6", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr6-{iso}.xls"),
                      ("coal", f"https://npp.gov.in/public-reports/cea/daily/fuel/{dd}/dailyCoal1-{iso}.xls")):
        r = get(url)
        if r is None or not r.ok:
            continue
        df = pd.read_excel(io.BytesIO(r.content), header=None)
        out(f"  {name} header rows (all columns):")
        for i in range(min(9, len(df))):
            out(f"   {i}: " + " | ".join(str(v).replace('\n', ' ') for v in df.iloc[i].tolist()))
        hits = df[df.apply(lambda r: r.astype(str).str.contains("ALL INDIA|All India|TOTAL", case=False).any(), axis=1)]
        out(f"  {name} total rows:\n" + hits.head(12).to_string(max_cols=40, max_colwidth=18)[:5000])
        if name == "dgr2":
            types = df[df[0].astype(str).str.strip().str.upper().eq("TYPE:")]
            out("  TYPE values: " + str(sorted(set(types.iloc[:, 1:8].astype(str).agg(" ".join, axis=1)))))
    for d in (date(2021, 1, 15), date(2019, 6, 15)):
        dd, iso = d.strftime("%d-%m-%Y"), d.isoformat()
        for name, url in (("dgr1", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr1-{iso}.xls"),
                          ("dgr2", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr2-{iso}.xls"),
                          ("coal", f"https://npp.gov.in/public-reports/cea/daily/fuel/{dd}/dailyCoal1-{iso}.xls")):
            get(url)


def cea():
    for y, m, d in ((2025, 8, 31), (2024, 3, 31), (2023, 3, 31), (2026, 7, 31)):
        for name in (f"IC_allocation_as_on_{d:02d}.{m:02d}.{y}.xlsx", f"IC_{date(y, m, 1):%b}_{y}.pdf"):
            get(f"https://cea.nic.in/wp-content/uploads/installed/{y}/{m:02d}/{name}")
    r = get("https://cea.nic.in/installed-capacity-report/?lang=en")
    if r is not None:
        out("  links: " + ", ".join(sorted(set(re.findall(r'(https://cea\.nic\.in/wp-content/uploads/installed/[^"\']+)',
                                                         r.text)))[:40]))


def igx():
    r = get("https://exchange.igxindia.com/api/gixi/get-gixi-cumulative-data")
    if r is not None:
        out(r.text[:1500])
    for q in ("?DeliveryMonth=Oct'26", "?DeliveryMonth=2026-10-01", "?DeliveryMonth=10-2026", "?DeliveryMonth=Oct-2026"):
        r = get("https://exchange.igxindia.com/api/gixi/get-gixi-total-monthly" + q)
        if r is not None:
            out(f"  {q}: {r.text[:400]}")


def pgcb():
    r = get("https://erp.powergrid.gov.bd/w/generations/view_generations", verify=False)
    if r is not None:
        i = r.text.find("<tbody")
        out(r.text[i:i + 3000])


if __name__ == "__main__":
    for f in sys.argv[1:] or ["pgcb", "ppac", "npp", "cea", "igx"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
