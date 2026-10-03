"""
South & Southeast Asia dashboard, discovery round 1: Bangladesh, Sri Lanka, Bhutan, Nepal,
Pakistan. Candidates from web research and open-source scrapers (electricitymaps ERP_PGCB.py,
demandcast pucsl.py, sasea-demand); all hosts are blocked from the Claude sandbox:

  Bangladesh  PGCB hourly generation by fuel (erp.powergrid.gov.bd HTML table), BPDB daily
              NLDC report archive (PDF), Petrobangla daily gas report listing, HCU monthly
  Sri Lanka   PUCSL gendata API - earlier probes (discovery_archive/rest_of_world/SRILANKA_*)
              got HTTP 500; scrapers pass dateAggregation + from/to (ISO, <= 7 days)
  Bhutan      BPSO publicdata/energy_data/{code}/{YYYY-MM-01} JSON (Cloudflare)
  Nepal       NEA Nepal Daily Operational Report (NDOR) PDFs, Bikram Sambat dates
  Pakistan    PMD Flood Forecasting Division station JSON (Tarbela / Mangla levels),
              NEPRA admission notices (monthly energy purchase PDFs), ISMO

Nothing is written.
"""
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone

import requests
import urllib3

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "asia"))
from SEA_DISCOVERY_TH_VN_PH import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()
TODAY = date.today()


def bangladesh():
    for u in ["https://erp.powergrid.gov.bd/w/generations/view_generations",
              "https://erp.powergrid.gov.bd/w/generations/view_generations?page=2000",
              "https://erp.pgcb.gov.bd/web/generations/view_generations"]:
        probe("PGCB", u, show=3000)
        probe("PGCB (verify off)", u, show=3000, verify=False)
    probe("BPDB daily archive", "https://misc.bpdb.gov.bd/daily-generation-archive", show=2000)
    probe("BPDB area demand", "https://misc.bpdb.gov.bd/area-wise-demand?date="
          + (TODAY - timedelta(days=2)).strftime("%d-%m-%Y"), show=1500)
    probe("Petrobangla reports (old portal)", "https://petrobangla.portal.gov.bd/site/view/reports", show=1500)
    probe("Petrobangla new site", "https://petrobangla.org.bd/", show=1500)
    probe("HCU monthly", "https://hcu.portal.gov.bd/site/page/monthly_report", show=1500)


def sri_lanka():
    to = datetime.combine(TODAY - timedelta(days=1), datetime.min.time()).replace(tzinfo=timezone.utc)
    for frm, too in [(to - timedelta(days=2), to), (datetime(2023, 1, 1, tzinfo=timezone.utc),
                                                    datetime(2023, 1, 3, tzinfo=timezone.utc))]:
        f, t = frm.strftime("%Y-%m-%dT%H:%M:%S.000Z"), too.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        for agg in ("15min", "1h", "1d"):
            probe(f"PUCSL dispatch {agg}", "https://gendata.pucsl.gov.lk/api/actual-system-dispatch"
                  f"?dateAggregation={agg}&from={f}&to={t}", show=1500)
        probe("PUCSL reservoir", f"https://gendata.pucsl.gov.lk/api/reservoir/storage-rainfall?from={f}&to={t}",
              show=800)
        probe("PUCSL daily stats", "https://gendata.pucsl.gov.lk/api/daily-generation-statistics"
              f"?from={f}&to={t}", show=800)
    for u in ("metadata/power-plants", "metadata/power-plant-complexes", "metadata/latest-year"):
        probe("PUCSL meta", f"https://gendata.pucsl.gov.lk/api/{u}", show=3000)


def bhutan():
    hdr = dict(H, Referer="https://www.bpso.bt/home/energy", **{"X-Requested-With": "XMLHttpRequest"})
    m = TODAY.replace(day=1) - timedelta(days=1)
    for code in ("generation_mwh", "peak_demand_mw", "energy_export_mwh"):
        for d in (m.replace(day=1), date(2023, 4, 1)):
            u = f"https://www.bpso.bt/publicdata/energy_data/{code}/{d.isoformat()}"
            out(f"\n{'=' * 90}\nBPSO {u}")
            try:
                r = requests.get(u, headers=hdr, timeout=T, verify=False)
                out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} {r.text[:800]}")
            except Exception as e:
                out(f"  ERROR {e!r}")


def nepal():
    r = probe("NEA NDOR listing", "https://www.nea.org.np/dailyOperationalReports", show=1500, links=40)
    if r is not None and r.ok:
        pdfs = re.findall(r"""["']([^"']*NDOR[^"']*\.pdf)["']""", r.text, re.I)
        out(f"  NDOR pdf links: {pdfs[:10]}")
    probe("NEA cms", "https://cms.nea.org.np/", show=500)


def pakistan():
    r = probe("FFD rivers", "https://ffd.pmd.gov.pk/flood-assistant/rivers", show=2000, links=80)
    probe("FFD river status", "https://ffd.pmd.gov.pk/home/river-status", show=2000)
    for sid in (1, 2, 3, 10):
        probe(f"FFD station {sid}", f"https://ffd.pmd.gov.pk/flood-assistant/station/{sid}?lang=en", show=1200)
    probe("NEPRA admission notices", "https://nepra.org.pk/Admission%20Notices/", show=1500, links=80)
    probe("NEPRA notices page", "https://nepra.org.pk/admission-notices.php", show=1500, links=80)
    probe("ISMO", "https://ismo.gov.pk/", show=1500)
    probe("NTDC", "https://ntdc.gov.pk/", show=500)
    probe("WAPDA river flow", "https://www.wapda.gov.pk/river-flow", show=1500)


if __name__ == "__main__":
    for f in sys.argv[1:] or ["bangladesh", "sri_lanka", "bhutan", "nepal", "pakistan"]:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
