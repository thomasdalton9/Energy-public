"""Probe 1 (12 Oct 2026): monthly installed capacity sources for DE, NL, BE, LU. Prints status/size/head of each candidate.

FINDINGS of probes 1-8 and DE_MASTR_DIAG.py (12 Oct 2026; outputs in discovery_archive/results/europe_capacity/de_nl_be_lu/):
 Germany  MaStR 'Gesamtdatenexport' zip (download.marktstammdatenregister.de/Gesamtdatenexport_<date>_<ver>.zip, link on the Datendownload page, 3.2 GB, UTF-16 XML,
          Range requests work; a single stream gives ~0.7 MB/s, 8 parallel ~6 MB/s) holds every unit with Inbetriebnahmedatum, DatumEndgueltigeStilllegung, status, net/gross kW:
          month-end capacity by fuel can be rebuilt exactly. Built (europe/capacity/GERMANY_MASTR_CAPACITY_MONTHLY.py). Energy-Charts monthly exists for DE only
          (renewables + batteries, equals MaStR). SMARD installed power is annual. The MaStR web API/overview is paged JSON per unit (not used).
 Belgium  Elia open data: ods031 (wind) / ods032 (PV) monitoredcapacity time series = the series ENTSO-E reports; ods036 (units by fuel, 1 Jan 2015-2023) and ods179
          (contracted units, daily from 21 May 2024). Built (BENELUX_CAPACITY.py). ODS date_format needs lower-case 'yyyy-MM-dd'; an aliased date_format() comes back under the expression text.
 Netherlands CBS StatLine 86266NED (MW by carrier), 82610NED (wind on land/sea, biomass), 85005NED (solar kWp) are ANNUAL; no monthly/quarterly source found
          (CertiQ, TenneT, energieopwek.nl, api.energieopwek.nl, rvo.nl: unreachable or no capacity series). Energy-Charts NL is yearly. Built annual only.
 Luxembourg ILR 'Donnees statistiques du secteur d'electricite' (data.public.lu) sheet '2. Production': installed kW by source 2022-2025 only. Nothing monthly (Creos, STATEC,
          data.public.lu searches). Energy-Charts LU is yearly. Built annual only."""
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
