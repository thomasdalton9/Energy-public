"""
South & Southeast Asia dashboard, discovery round 1: India. Candidates from web research and
open-source scrapers (electricitymaps parsers/IN.py, Grid-Sentinel); every .gov.in host is
blocked from the Claude sandbox and some are said to be geofenced to Indian IPs, so this
checks what a GitHub Actions runner can reach:

  Grid-India daily PSP report (generation by fuel, peak / energy met): POST
      webapi.grid-india.in/api/v1/file {"_source":"GRDW","_type":"DAILY_PSP_REPORT",...},
      files at webcdn.grid-india.in/<FilePath>
  NPP (CEA) date-addressed reports: dgr2 (thermal/nuclear/hydro generation), dgr6 (hydro
      reservoirs), dailyCoal1 (coal stocks), installcap (monthly installed capacity)
  CEA daily RE generation (admin-ajax JSON)
  CWC reservoir bulletin (RSMS)
  PPAC gas production / consumption (AjaxController POST), sectoral consumption + LNG import xlsx
  IEX day-ahead market snapshot JSON

Downloads one sample of each file found and prints its sheets / first rows. Nothing is written.
"""
import io
import json
import os
import sys
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "asia"))
from SEA_DISCOVERY_TH_VN_PH import H, T, out, probe  # noqa: E402

TODAY = date.today()


def show_excel(content, label):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        out(f"  {label}: not excel ({e!r}); head {content[:200]!r}")
        return
    out(f"  {label}: sheets {xl.sheet_names}")
    for s in xl.sheet_names[:6]:
        df = pd.read_excel(xl, s, header=None)
        out(f"  --- sheet {s!r} shape {df.shape}")
        out(df.head(45).to_string(max_cols=14, max_colwidth=28)[:5000])


def fy(d):
    return f"{d.year}-{str(d.year + 1)[2:]}" if d.month >= 4 else f"{d.year - 1}-{str(d.year)[2:]}"


def grid_india():
    probe("Grid-India PSP page", "https://grid-india.in/en/reports/daily-psp-report", show=600)
    d = TODAY - timedelta(days=3)
    for fdate in (fy(d), f"{d.year if d.month >= 4 else d.year - 1}-{(d.year + 1) if d.month >= 4 else d.year}"):
        body = {"_source": "GRDW", "_type": "DAILY_PSP_REPORT", "_fileDate": fdate, "_month": f"{d.month:02d}"}
        out(f"\n{'=' * 90}\nGrid-India webapi POST {body}")
        try:
            r = requests.post("https://webapi.grid-india.in/api/v1/file", json=body,
                              headers=dict(H, Origin="https://grid-india.in", Referer="https://grid-india.in/"),
                              timeout=T)
            out(f"  -> {r.status_code} {len(r.content)} {r.text[:600]}")
            files = (r.json() or {}).get("retData") or []
        except Exception as e:
            out(f"  ERROR {e!r}")
            continue
        out(f"  {len(files)} files; sample {files[:3]}")
        xls = [f for f in files if str(f.get("FilePath", "")).lower().endswith((".xls", ".xlsx"))]
        if xls:
            f = xls[0]
            r = requests.get(f"https://webcdn.grid-india.in/{f['FilePath']}", headers=H, timeout=T)
            out(f"  file {f['FilePath']} -> {r.status_code} {len(r.content)}")
            if r.ok:
                show_excel(r.content, f["FilePath"])
            break
    # history: an old month
    body = {"_source": "GRDW", "_type": "DAILY_PSP_REPORT", "_fileDate": "2021-22", "_month": "06"}
    try:
        r = requests.post("https://webapi.grid-india.in/api/v1/file", json=body, headers=H, timeout=T)
        files = (r.json() or {}).get("retData") or []
        out(f"\nGrid-India 2021-06: {r.status_code} {len(files)} files; sample {files[:2]}")
    except Exception as e:
        out(f"  history ERROR {e!r}")


def npp():
    for back in (2, 3, 400):
        d = TODAY - timedelta(days=back)
        dd, iso = d.strftime("%d-%m-%Y"), d.isoformat()
        for name, url in [("dgr2", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr2-{iso}.xls"),
                          ("dgr6", f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr6-{iso}.xls"),
                          ("coal", f"https://npp.gov.in/public-reports/cea/daily/fuel/{dd}/dailyCoal1-{iso}.xlsx"),
                          ("coal xls", f"https://npp.gov.in/public-reports/cea/daily/fuel/{dd}/dailyCoal1-{iso}.xls")]:
            r = probe(f"NPP {name} {iso}", url, show=0)
            if r is not None and r.ok and back == 3:
                show_excel(r.content, name)
    m = (TODAY.replace(day=1) - timedelta(days=40))
    mon = m.strftime("%b").upper()
    for name in ("capacity1", "installcap1"):
        probe(f"NPP {name}", f"https://npp.gov.in/public-reports/cea/monthly/installcap/{m.year}/{mon}/"
              f"{name}-{m.year}-{m.month:02d}.pdf", show=0)
    probe("CEA installed capacity page", "https://cea.nic.in/installed-capacity-report/?lang=en", show=800)
    probe("CEA RE generation ajax", "https://cea.nic.in/wp-admin/admin-ajax.php?action=getpostsfordatatables"
          "&code=renewable", show=1500)


def cwc():
    probe("CWC RSMS bulletins", "https://rsms.cwc.gov.in/", show=1500)
    probe("CWC reservoir page", "https://cwc.gov.in/reservoir-level-storage-bulletin", show=1500)


def ppac():
    probe("PPAC production page", "https://ppac.gov.in/natural-gas/production", show=800)
    for path, page in (("getGasProduction", 170), ("getGasConsumption", 138)):
        out(f"\n{'=' * 90}\nPPAC POST {path}")
        try:
            s = requests.Session()
            s.get("https://ppac.gov.in/natural-gas/production", headers=H, timeout=T)
            r = s.post(f"https://ppac.gov.in/AjaxController/{path}",
                       data={"financialYear": fy(TODAY).replace("-", "-20") if False else
                             f"{fy(TODAY)[:4]}-{int(fy(TODAY)[:4]) + 1}", "reportBy": 4, "pageId": page},
                       headers=dict(H, **{"X-Requested-With": "XMLHttpRequest",
                                         "Referer": "https://ppac.gov.in/natural-gas/production"}), timeout=T)
            out(f"  -> {r.status_code} {r.headers.get('content-type')} {r.text[:3000]}")
        except Exception as e:
            out(f"  ERROR {e!r}")
    for pg in ("sectoral-consumption", "import"):
        probe(f"PPAC {pg}", f"https://ppac.gov.in/natural-gas/{pg}", show=600)


def iex():
    d = (TODAY - timedelta(days=1)).strftime("%d-%m-%Y")
    d0 = (TODAY - timedelta(days=3)).strftime("%d-%m-%Y")
    for u in [f"https://www.iexindia.com/api/v1/dam/market-snapshot?interval=ONE_HOUR&fromDate={d0}&toDate={d}",
              f"https://www.iexindia.com/api/v1/dam/area-price?interval=ONE_HOUR&fromDate={d0}&toDate={d}",
              "https://www.iexindia.com/api/v1/dam/market-snapshot?interval=ONE_HOUR&fromDate=01-01-2022&toDate=02-01-2022"]:
        probe("IEX", u, show=2000)
    probe("IGX GIXI", "https://exchange.igxindia.com/web/gixi", show=800)


if __name__ == "__main__":
    for f in sys.argv[1:] or ["grid_india", "npp", "cwc", "ppac", "iex"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
