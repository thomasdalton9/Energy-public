"""
South Asia discovery, round 2 (after SEA_DISCOVERY_INDIA / SEA_DISCOVERY_SOUTH_ASIA, which found from a
GitHub Actions runner: Grid-India hosts reset the connection (blocked); NPP date-addressed xls work
(dgr2 = plant-level thermal generation, dgr6 = hydro reservoirs, dailyCoal1 = coal stocks); CEA
admin-ajax lists daily RE generation xlsx; CEA installed capacity xlsx (IC_allocation_as_on_*.xlsx);
PPAC page xlsx links; IEX DAM JSON (history to 2022); IGX GIXI JSON paths; PGCB hourly table (verify
off) back to 2015; NEA behind a JS challenge; ISMO a JS app; FFD / NEPRA 403).

This round:
  India   NPP dgr1, dgr3-dgr12 (which holds the all-India generation summary by category?); one CEA
          daily RE xlsx; CEA IC xlsx; PPAC NG production / consumption / sectoral / LNG import xlsx;
          IGX GIXI JSON; IEX area-price paths; Grid-India alternative hosts
  Bangladesh  PGCB rows per page and a date search parameter
  Pakistan    ISMO app's JS bundle (API paths), NTDC pages
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
from SEA_DISCOVERY2 import get, show_excel  # noqa: E402

urllib3.disable_warnings()
D = date.today() - timedelta(days=3)


def npp():
    dd, iso = D.strftime("%d-%m-%Y"), D.isoformat()
    for k in [1, 3, 4, 5, 7, 8, 9, 10, 11, 12]:
        r = get(f"https://npp.gov.in/public-reports/cea/daily/dgr/{dd}/dgr{k}-{iso}.xls")
        if r is not None and r.ok:
            show_excel(r.content, f"dgr{k}", rows=40, sheets=1)


def cea():
    r = get("https://cea.nic.in/wp-admin/admin-ajax.php?action=getpostsfordatatables&code=renewable")
    if r is not None and r.ok:
        rows = r.json().get("data", [])
        out(f"  {len(rows)} RE rows; dates {rows[-1].get('date') if rows else ''}..{rows[0].get('date') if rows else ''}")
        out(f"  sample keys {list(rows[0].keys()) if rows else ''}")
        dates = sorted(x.get("date") for x in rows)
        out(f"  sorted dates {dates[:3]} .. {dates[-3:]}")
        latest = max(rows, key=lambda x: x.get("date", ""))
        x = latest.get("link", "").split("^")[0]
        r2 = get(x)
        if r2 is not None and r2.ok:
            show_excel(r2.content, "CEA RE", rows=45, sheets=2)
    for code in ("generation", "daily", "dgr"):
        r = get(f"https://cea.nic.in/wp-admin/admin-ajax.php?action=getpostsfordatatables&code={code}")
        if r is not None:
            out(f"  code={code}: {r.text[:500]}")
    r = get("https://cea.nic.in/wp-content/uploads/installed/2026/08/IC_allocation_as_on_31.08.2026.xlsx")
    if r is not None and r.ok:
        show_excel(r.content, "CEA IC", rows=30, sheets=2)


def ppac():
    for u in ("https://ppac.gov.in/download.php?file=reports/1790849513_1_NG-C-Production.xlsx",
              "https://ppac.gov.in/uploads/page-images/1790853468_4_NG-C-Sectoral-Consumption.xlsx",
              "https://ppac.gov.in/natural-gas/1781859871_2_NG-H_Sectoral_Consumption.xlsx",
              "https://ppac.gov.in/uploads/page-images/1790853207_2_NG-C-LNG-Import.xls",
              "https://ppac.gov.in/uploads/page-images/1781164918_NG-H_LNG_Import.xlsx"):
        r = get(u)
        if r is not None and r.ok:
            show_excel(r.content, u.rsplit("/", 1)[-1], rows=35, sheets=2)
    r = get("https://ppac.gov.in/natural-gas/consumption")
    if r is not None:
        out("  xlsx: " + ", ".join(sorted(set(re.findall(r'["\']([^"\']+\.xlsx?)["\']', r.text)))))


def igx():
    for u in ("https://exchange.igxindia.com/api/gixi/get-gixi-cumulative-data",
              "https://exchange.igxindia.com/api/gixi/get-gixi-total-monthly?DeliveryMonth=",
              "https://exchange.igxindia.com/api/gixi/get-gixi-total-monthly?DeliveryMonth=2024-01"):
        r = get(u)
        if r is not None:
            out(f"  {r.text[:1500]}")
    r = get("https://exchange.igxindia.com/web/gixi")
    if r is not None:
        out("  api paths: " + ", ".join(sorted(set(re.findall(r'(/api/[\w/-]+)', r.text)))))


def iex():
    d1, d0 = (date.today() - timedelta(days=1)).strftime("%d-%m-%Y"), (date.today() - timedelta(days=2)).strftime("%d-%m-%Y")
    for path in ("dam/area-prices", "dam/area-wise-price", "dam/area-price-snapshot", "rtm/market-snapshot",
                 "gdam/market-snapshot"):
        r = get(f"https://www.iexindia.com/api/v1/{path}?interval=ONE_HOUR&fromDate={d0}&toDate={d1}")
        if r is not None:
            out(f"  {r.text[:300]}")
    r = get("https://www.iexindia.com/api/v1/dam/market-snapshot?interval=ONE_HOUR&fromDate=01-01-2020&toDate=31-01-2020")
    if r is not None:
        out(f"  2020: {r.text[:300]}")
    r = get("https://www.iexindia.com/api/v1/dam/market-snapshot?interval=ONE_HOUR&fromDate=01-08-2026&toDate=31-08-2026")
    if r is not None:
        out(f"  31-day range: n={len((r.json() or {}).get('data') or []) if r.ok else r.text[:200]}")


def grid_india():
    for u in ("http://grid-india.in/", "https://posoco.in/", "https://report.grid-india.in/",
              "https://webcdn.grid-india.in/"):
        get(u)


def pgcb():
    for u in ("https://erp.powergrid.gov.bd/w/generations/view_generations?search=01-09-2026",
              "https://erp.powergrid.gov.bd/w/generations/view_generations?date=01-09-2026",
              "https://erp.powergrid.gov.bd/w/generations/view_generations?page=2"):
        r = get(u, verify=False)
        if r is not None and r.ok:
            rows = re.findall(r"\d\d-\d\d-\d{4} \d\d:\d\d:\d\d", r.text)
            out(f"  {len(rows)} timestamps: {rows[:2]} .. {rows[-2:]}")
            forms = re.findall(r"(?is)<form.*?</form>", r.text)
            out("  forms: " + str([re.sub(r"\s+", " ", f)[:400] for f in forms[:2]]))
            out("  pagination: " + ", ".join(sorted(set(re.findall(r'href="([^"]*page=\d+[^"]*)"', r.text)))[:8]))


def pakistan():
    r = get("https://ismo.gov.pk/")
    if r is not None:
        for js in re.findall(r'src="([^"]+\.js)"', r.text)[:6]:
            j = get(js if js.startswith("http") else "https://ismo.gov.pk" + ("" if js.startswith("/") else "/") + js)
            if j is not None:
                paths = sorted(set(re.findall(r'["\'`](https?://[^"\'`]*api[^"\'`]*|/api/[^"\'`]+)', j.text)))
                out(f"  api paths: {paths[:60]}")
                out("  keywords: " + ", ".join(sorted(set(re.findall(r'(generation|demand|dispatch|load|report)[\w-]*',
                                                                        j.text, re.I)))[:60]))
    for u in ("https://ntdc.gov.pk/reports", "https://ntdc.gov.pk/energy-data", "https://ntdc.gov.pk/power-data"):
        r = get(u)
        if r is not None and r.ok:
            out(f"  {re.sub(r'<[^>]+>', ' ', r.text)[:600]}")


if __name__ == "__main__":
    for f in sys.argv[1:] or ["npp", "cea", "ppac", "igx", "iex", "grid_india", "pgcb", "pakistan"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
