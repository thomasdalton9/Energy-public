"""
Round 2 probe of national TSO open-data endpoints for the Europe dashboard (keyless only):
exact dataset ids / fields for generation by fuel, installed capacity and hydro reservoirs.
Manual/push-only; archive when settled.
"""
import json

import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}


def get(url, **params):
    try:
        r = requests.get(url, headers=H, params=params, timeout=(15, 90))
        return r
    except requests.RequestException as e:
        print("  ERROR", type(e).__name__, str(e)[:200])


def show(label, url, n=1200, **params):
    print(f"\n=== {label}: {url} {params}", flush=True)
    r = get(url, **params)
    if r is None:
        return None
    print("  status", r.status_code, "bytes", len(r.content))
    print("  body:", r.text[:n].replace("\n", " "))
    return r


# Elia: list every dataset title (108) so generation / capacity ones can be picked
r = get("https://opendata.elia.be/api/explore/v2.1/catalog/datasets", limit=100, select="dataset_id,title")
if r is not None and r.ok:
    print("\n=== Elia datasets")
    for d in r.json()["results"]:
        print(" ", d["dataset_id"], "|", d["title"])
    r = get("https://opendata.elia.be/api/explore/v2.1/catalog/datasets", limit=100, offset=100, select="dataset_id,title")
    for d in (r.json()["results"] if r is not None and r.ok else []):
        print(" ", d["dataset_id"], "|", d["title"])

# NESO: the historic generation mix resource
r = show("NESO historic generation mix sample", "https://api.neso.energy/api/3/action/datastore_search", n=1800,
         resource_id="f93d1835-75bc-43e5-84ad-12472b180a98", limit=2)
show("NESO installed capacity / other packages", "https://api.neso.energy/api/3/action/package_search", n=200, q="capacity", rows=1)
r = get("https://api.neso.energy/api/3/action/package_search", q="generation", rows=15)
if r is not None and r.ok:
    print("\n=== NESO packages for 'generation'")
    for p in r.json()["result"]["results"]:
        print(" ", p["name"], "|", p["title"], "|", [(x["id"], x["name"], x.get("format")) for x in p["resources"]][:4])

# Spain REE: other hydro / capacity widgets
for widget in ("generacion/potencia-instalada", "generacion/estructura-renovables", "balance/balance-electrico",
               "generacion/evolucion-renovable-no-renovable"):
    show("REE " + widget, f"https://apidatos.ree.es/en/datos/{widget}", n=500,
         start_date="2025-01-01T00:00", end_date="2025-03-31T23:59", time_trunc="month")
show("REE reserva-hidraulica (month)", "https://apidatos.ree.es/en/datos/hidraulica/reserva-hidraulica", n=500,
     start_date="2025-01-01T00:00", end_date="2025-03-31T23:59", time_trunc="day")
show("REE reserva-hidraulica (week-ish, es)", "https://apidatos.ree.es/es/datos/hidraulica/reserva-hidraulica", n=500,
     start_date="2025-01-01T00:00", end_date="2025-03-31T23:59", time_trunc="day")

# Statnett: field names
r = get("https://driftsdata.statnett.no/restapi/ProductionConsumption/GetData", From="2025-01-01")
if r is not None and r.ok:
    j = r.json()
    print("\n=== Statnett keys:", {k: (len(v) if isinstance(v, list) else v) for k, v in j.items()})
show("Statnett production by type", "https://driftsdata.statnett.no/restapi/Production/GetData", n=600, From="2025-01-01")
show("Statnett ProductionConsumption lastest", "https://driftsdata.statnett.no/restapi/ProductionConsumption/GetLatestDetailedOverview", n=1500)

# Poland PSE: metadata for generation endpoints
show("PSE metadata", "https://api.raporty.pse.pl/api/$metadata", n=1800)
show("PSE gen-jw sample", "https://api.raporty.pse.pl/api/gen-jw", n=800, **{"$first": 2})
show("PSE his-gen-paliwo?", "https://api.raporty.pse.pl/api/gen-paliwo", n=800, **{"$first": 2})

# Energinet datasets with capacity / generation
r = get("https://api.energidataservice.dk/meta/dataset")
if r is not None and r.ok:
    print("\n=== Energinet datasets (name | title)")
    for d in r.json():
        if d.get("organizationName", "").startswith("tso-"):
            print(" ", d["datasetName"], "|", d["title"][:90], "|", d["organizationName"])
