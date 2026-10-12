"""Probe 1 (12 Oct 2026): monthly installed capacity sources for DE, NL, BE, LU. Prints status/size/head of each candidate.
Findings: see discovery_archive/results/europe_capacity/ after running."""
import json, sys, requests
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"
def get(u, n=500, **kw):
    try:
        r = S.get(u, timeout=60, **kw)
        print(f"\n== {u}\n   {r.status_code} {r.headers.get('content-type')} len={len(r.content)} lm={r.headers.get('last-modified')}")
        print("  ", r.text[:n].replace("\n", " ") if "text" in (r.headers.get("content-type") or "") or "json" in (r.headers.get("content-type") or "") else "")
        return r
    except Exception as e:
        print(f"\n== {u}\n   ERR {e}")
for u in [
 "https://www.marktstammdatenregister.de/MaStR/Datendownload",
 "https://www.marktstammdatenregister.de/MaStRHilfe/subpages/statistik.html",
 "https://www.marktstammdatenregister.de/MaStR/Einheit/Einheiten/OeffentlicheEinheitenuebersicht",
 "https://api.energy-charts.info/installed_power?country=de&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=nl&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=be&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=lu&time_step=monthly&installation_decommission=false",
 "https://www.smard.de/app/chart_data/1004071/DE/index_year.json",
 "https://www.smard.de/nip-download-manager/nip/download/market-data",
 "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?search=installed&limit=50&select=dataset_id,title",
 "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?search=capacity&limit=50&select=dataset_id,title",
 "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?search=photovoltaic&limit=50&select=dataset_id,title",
 "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?search=wind&limit=50&select=dataset_id,title",
 "https://odata4.cbs.nl/CBS/85005NED/Properties",
 "https://opendata.cbs.nl/ODataCatalog/Tables?$format=json&$filter=substringof('vermogen',Title)",
 "https://opendata.cbs.nl/ODataCatalog/Tables?$format=json&$filter=substringof('Zonnestroom',Title)",
 "https://opendata.cbs.nl/ODataCatalog/Tables?$format=json&$filter=substringof('Windenergie',Title)",
 "https://opendata.cbs.nl/ODataCatalog/Tables?$format=json&$filter=substringof('lektriciteit',Title)",
 "https://data.public.lu/api/1/datasets/?q=electricity+capacity&page_size=20",
 "https://data.public.lu/api/1/datasets/?q=photovoltaique&page_size=20",
 "https://data.public.lu/api/1/datasets/?q=production+electricite&page_size=20",
]:
    r = get(u, 1500)
    if r is not None and "elia" in u and r.ok:
        pass
