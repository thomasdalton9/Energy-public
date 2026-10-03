"""
South & Southeast Asia dashboard, discovery round 1: Malaysia, Indonesia, Brunei.
Candidates from web research (sites blocked from the Claude sandbox), checked from a
GitHub Actions runner:

  Malaysia   GSO (Grid System Operator) page methods - POST JSON {"Fromdate","Todate"}:
               SystemData/CurrentGen.aspx/GetChartDataSource (half-hourly generation by fuel),
               SystemData/SystemDemand.aspx/GetChartDataSource, PowerStation.aspx/GetDataSource;
             Single Buyer API (system marginal price, generation mix);
             data.gov.my / DOSM electricity_supply / electricity_consumption csv (is it still updated?)
  Indonesia  ESDM HEESI handbook (pdf/xlsx), BPS public pages (the web API needs a key)
  Brunei     DEPS eData library (gas production, LNG exports, LNG price)

Prints status, size, a snippet and data-like links. For GSO it asks for a recent week and a
date years back to see how far history goes. Nothing is written.
"""
import io
import json
import re
import sys
from datetime import date, timedelta

import requests

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from SEA_DISCOVERY_TH_VN_PH import H, T, out, probe  # noqa: E402

GSO = "https://www.gso.org.my/SystemData"


def gso_post(page, method, d0, d1):
    url = f"{GSO}/{page}/{method}"
    body = {"Fromdate": d0.strftime("%d/%m/%Y"), "Todate": d1.strftime("%d/%m/%Y")}
    hdr = dict(H, **{"Content-Type": "application/json; charset=utf-8", "Origin": "https://www.gso.org.my",
                     "Referer": f"{GSO}/{page}", "X-Requested-With": "XMLHttpRequest"})
    out(f"\n{'=' * 90}\nGSO POST {url} {body}")
    try:
        r = requests.post(url, headers=hdr, data=json.dumps(body), timeout=T)
    except Exception as e:
        out(f"  ERROR {e!r}")
        return
    out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
    try:
        j = r.json()
        d = j.get("d", j)
        if isinstance(d, str):
            d = json.loads(d)
        if isinstance(d, list):
            out(f"  {len(d)} records; first {d[:2]}; last {d[-1:]}")
            if d and isinstance(d[0], dict):
                out(f"  keys {list(d[0].keys())}")
        else:
            out(f"  {str(d)[:1500]}")
    except Exception:
        out(f"  text {r.text[:1500]}")


def malaysia():
    probe("GSO home", "https://www.gso.org.my/SystemData/CurrentGen.aspx", show=1500)
    t = date.today()
    for page, method in [("CurrentGen.aspx", "GetChartDataSource"), ("SystemDemand.aspx", "GetChartDataSource"),
                         ("FuelMix.aspx", "GetChartDataSource"), ("PowerStation.aspx", "GetDataSource"),
                         ("TieLine.aspx", "GetChartDataSource")]:
        gso_post(page, method, t - timedelta(days=2), t - timedelta(days=1))
    for back in (30, 365, 3 * 365, 6 * 365):   # how far back does history go?
        d = t - timedelta(days=back)
        gso_post("CurrentGen.aspx", "GetChartDataSource", d, d)
    probe("Single Buyer SMP page", "https://www.singlebuyer.com.my/resources-marginal.php")
    probe("Single Buyer SMP api", "https://www.singlebuyer.com.my/api/v1/smp/actual-forecast")
    probe("Single Buyer gen mix api", "https://www.singlebuyer.com.my/api/v1/charts/op-scheduling/day-generation-mix")
    for ds in ("electricity_supply", "electricity_consumption"):
        r = probe(f"data.gov.my {ds}", f"https://storage.data.gov.my/energy/{ds}.csv", show=0)
        if r is not None and r.ok:
            lines = r.text.strip().splitlines()
            out(f"  {len(lines)} lines; head {lines[:3]}; tail {lines[-3:]}")
    probe("data.gov.my catalogue", "https://api.data.gov.my/data-catalogue?id=electricity_supply&limit=3&sort=-date")
    probe("MEIH", "https://meih.st.gov.my/statistics")
    probe("MyEnergyStats", "https://myenergystats.st.gov.my/")


def indonesia():
    probe("HEESI landing", "https://esdm.go.id/en/publikasi/handbook-of-energy-economic-statistics-of-indonesia-heesi")
    for y in (2025, 2024):
        probe(f"HEESI {y} pdf", "https://esdm.go.id/assets/media/content/"
              f"content-handbook-of-energy-and-economic-statistics-of-indonesia-{y}.pdf", show=0)
    probe("Gatrik statistik", "https://gatrik.esdm.go.id/frontend/download_index?kode_category=statistik")
    probe("BPS LNG exports table", "https://www.bps.go.id/id/statistics-table/1/MTAxMyMx")
    probe("BPS webapi (no key)", "https://webapi.bps.go.id/v1/api/list/model/subject/lang/eng/domain/0000/")
    probe("SKK Migas", "https://www.skkmigas.go.id/")
    probe("Ditjen Migas", "https://migas.esdm.go.id/")


def brunei():
    probe("DEPS eData", "https://deps.mofe.gov.bn/?p=1665", show=3000, links=120)
    probe("DEPS home", "https://deps.mofe.gov.bn/")


if __name__ == "__main__":
    for f in sys.argv[1:] or ["malaysia", "indonesia", "brunei"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
