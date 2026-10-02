"""
Probe of national TSO open-data endpoints for the Europe dashboard (keyless ones only).
Prints status, field names and sample rows per endpoint - read the log. Manual-only; archive when settled.
"""
import json

import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}


def show(label, url, **params):
    print(f"\n=== {label}: {url} {params}", flush=True)
    try:
        r = requests.get(url, headers=H, params=params, timeout=(15, 90))
    except requests.RequestException as e:
        print("  ERROR", type(e).__name__, str(e)[:200])
        return
    print("  status", r.status_code, "bytes", len(r.content), r.headers.get("content-type"))
    print("  body:", r.text[:900].replace("\n", " "))


# Belgium - Elia open data (Opendatasoft)
show("Elia catalog: generation", "https://opendata.elia.be/api/explore/v2.1/catalog/datasets",
     search="generation", limit=15, select="dataset_id,title")
show("Elia catalog: installed capacity", "https://opendata.elia.be/api/explore/v2.1/catalog/datasets",
     search="installed capacity", limit=10, select="dataset_id,title")
# Denmark - Energinet
show("Energinet prod+cons settlement", "https://api.energidataservice.dk/dataset/ProductionConsumptionSettlement",
     limit=2, start="2025-01-01", end="2025-01-02")
show("Energinet datasets", "https://api.energidataservice.dk/meta/dataset")
# GB - NESO
show("NESO package search", "https://api.neso.energy/api/3/action/package_search", q="historic generation mix", rows=5)
# Spain - REE
show("REE generation structure", "https://apidatos.ree.es/en/datos/generacion/estructura-generacion",
     start_date="2025-01-01T00:00", end_date="2025-01-31T23:59", time_trunc="day")
show("REE reservoirs", "https://apidatos.ree.es/en/datos/generacion/evolucion-renovable-no-renovable",
     start_date="2025-01-01T00:00", end_date="2025-01-31T23:59", time_trunc="day")
show("REE hydro reservoirs", "https://apidatos.ree.es/en/datos/hidraulica/reserva-hidraulica",
     start_date="2025-01-01T00:00", end_date="2025-01-31T23:59", time_trunc="day")
# Norway - Statnett
show("Statnett prod/cons", "https://driftsdata.statnett.no/restapi/ProductionConsumption/GetData", From="2025-01-01")
# Poland - PSE
show("PSE gen by fuel", "https://api.raporty.pse.pl/api/gen-jw", **{"$filter": "doba eq '2025-01-10'", "$top": 3})
show("PSE his-wlk-cal", "https://api.raporty.pse.pl/api/his-wlk-cal", **{"$filter": "business_date eq '2025-01-10'", "$top": 3})
# Austria - APG
show("APG generation", "https://transparency.apg.at/api/v1/Data/AGPT/English/M15/2025-01-10T000000/2025-01-11T000000")
# Ireland - EirGrid smartgrid
show("EirGrid", "https://www.smartgriddashboard.com/DashboardService.svc/data",
     area="generationactual", region="ALL", datefrom="10-Jan-2025 00:00", dateto="11-Jan-2025 00:00")
# Czech - CEPS
show("CEPS", "https://www.ceps.cz/en/all-data")
# Switzerland - Swissgrid
show("Swissgrid", "https://www.swissgrid.ch/en/home/operation/grid-data/generation.html")
# Sweden - Svenska kraftnat
show("SvK", "https://www.svk.se/en/national-grid/the-control-room/")
# Gas: GRTgaz open data
show("GRTgaz catalog", "https://opendata.grtgaz.com/api/explore/v2.1/catalog/datasets", limit=30, select="dataset_id,title")
# Gas: Energinet gas
show("Energinet gas", "https://api.energidataservice.dk/dataset/GasFlowDK", limit=2)
# probe run 1
