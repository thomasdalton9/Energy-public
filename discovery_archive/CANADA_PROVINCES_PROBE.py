"""Probe Canadian provincial grid-operator / CER / StatCan endpoints (discovery only)."""
import re
import sys
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
URLS = [
    # AESO
    "http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet",
    "https://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet",
    "https://api.aeso.ca/report/v1.1/csd/summary/current",
    "https://www.aeso.ca/market/market-and-system-reporting/data-requests/",
    "https://www.aeso.ca/market/market-and-system-reporting/data-requests/hourly-metered-volumes-and-pool-price-and-ail-data/",
    "https://www.aeso.ca/market/market-and-system-reporting/electricity-statistics/",
    "http://ets.aeso.ca/ets_web/ip/Market/Reports/ActualForecastWMRQHReportServlet",
    "https://www.aeso.ca/grid/grid-planning/forecasting/long-term-adequacy-metrics/",
    # Hydro-Quebec
    "https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets?limit=100",
    "https://www.hydroquebec.com/data/documents-donnees/donnees-ouvertes/json/demande.json",
    "https://www.hydroquebec.com/data/documents-donnees/donnees-ouvertes/json/production.json",
    "https://www.hydroquebec.com/documents-donnees/donnees-ouvertes/",
    # BC Hydro
    "https://www.bchydro.com/energy-in-bc/operations/transmission/transmission-system/balancing-authority-load-data.html",
    "https://www.bchydro.com/energy-in-bc/operations/transmission/transmission-system/actual-flow-data.html",
    "https://www.bchydro.com/energy-in-bc/operations/transmission/transmission-system/balancing-authority-load-data/historical-transmission-data.html",
    # SaskPower
    "https://www.saskpower.com/about-us/our-company/blog/saskpower-system-data",
    "https://www.saskpower.com/about-us/our-company/power-system/power-generation",
    "https://data.saskatchewan.ca/api/3/action/package_search?q=saskpower",
    "https://www.saskpower.com/our-power-future/our-electricity-supply-and-demand",
    # NB Power
    "https://tso.nbpower.com/Public/en/system_information_archive.aspx",
    "https://tso.nbpower.com/Public/en/system_information.aspx",
    "https://tso.nbpower.com/Public/en/op/market/hourly_load.aspx",
    # Nova Scotia
    "https://www.nspower.ca/about-us/electricity/system-data",
    "https://www.nspower.ca/oasis/system-reports",
    "https://www.nspower.ca/about-us/electricity/system-data/hourly-and-forecast-load",
    "https://www.nspower.ca/clean-electricity/system-data",
    # Manitoba
    "https://www.hydro.mb.ca/corporate/water_regimes/",
    "https://www.hydro.mb.ca/corporate/water_regimes/water_levels.shtml",
    "https://www.hydro.mb.ca/corporate/water_regimes/current_data.shtml",
    "https://www.hydro.mb.ca/regulatory_affairs/",
    # CER
    "https://www.cer-rec.gc.ca/open/energy/",
    "https://www.cer-rec.gc.ca/open/imports-exports/natural-gas-exports-and-imports-by-month.csv",
    "https://www.cer-rec.gc.ca/open/energy/throughput/pipelines.csv",
    "https://www.cer-rec.gc.ca/open/energy/energyfutures2023/results-resultats.csv",
    "https://www.cer-rec.gc.ca/open/energy/energyfutures2026/",
    "https://www.cer-rec.gc.ca/en/data-analysis/energy-markets/provincial-territorial-energy-profiles/index.html",
    "https://open.canada.ca/data/api/3/action/package_search?q=energy+futures&rows=20",
    "https://open.canada.ca/data/api/3/action/package_search?q=natural+gas+exports+imports&rows=10",
    "https://open.canada.ca/data/api/3/action/package_search?q=reservoir+storage+hydro&rows=10",
    # StatCan WDS
    "https://www150.statcan.gc.ca/t1/wds/rest/getCubeMetadata",
    "https://www150.statcan.gc.ca/n1/tbl/csv/25100084-eng.zip",
    "https://www150.statcan.gc.ca/n1/tbl/csv/25100020-eng.zip",
]


def show(u):
    try:
        r = requests.get(u, headers=H, timeout=(10, 40))
    except Exception as e:  # noqa: BLE001
        print(f"ERR  {u}  {type(e).__name__}: {str(e)[:120]}", flush=True)
        return
    t = r.text[:300].replace("\n", " ") if "text" in r.headers.get("content-type", "") or "json" in r.headers.get("content-type", "") else "<binary>"
    print(f"{r.status_code} {u} ct={r.headers.get('content-type')} len={len(r.content)}\n     {t}", flush=True)
    if r.status_code == 200 and "html" in r.headers.get("content-type", ""):
        links = set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|zip|json|xml)[^"]*)"', r.text, re.I))
        for l in sorted(links)[:40]:
            print("     LINK", l)
    if r.status_code == 200 and "json" in r.headers.get("content-type", "") and "hydroquebec" in u and "catalog" in u:
        for d in r.json().get("results", []):
            print("     DS", d.get("dataset_id"), "|", (d.get("metas", {}).get("default", {}) or {}).get("title"))


for u in URLS:
    show(u)
