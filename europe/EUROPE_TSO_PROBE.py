"""Round 3 probe (keyless): Elia generation/capacity datasets, Energinet generation + gas samples. Push-triggered; archive when settled."""
import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}


def get(url, **params):
    try:
        return requests.get(url, headers=H, params=params, timeout=(15, 90))
    except requests.RequestException as e:
        print("  ERROR", type(e).__name__, str(e)[:200])


def show(label, url, n=1500, **params):
    print(f"\n=== {label}: {url} {params}", flush=True)
    r = get(url, **params)
    if r is not None:
        print("  status", r.status_code, "bytes", len(r.content))
        print("  body:", r.text[:n].replace("\n", " "))


KEY = ("generat", "fuel", "capacity", "installed", "production", "wind", "photovolt", "solar", "hydro", "reservoir")
for off in (0, 100):
    r = get("https://opendata.elia.be/api/explore/v2.1/catalog/datasets", limit=100, offset=off, select="dataset_id,title")
    for d in (r.json()["results"] if r is not None and r.ok else []):
        if any(k in d["title"].lower() for k in KEY):
            print("ELIA", d["dataset_id"], "|", d["title"])

for ds in ("ods201", "ods177", "ods033", "ods032"):
    show(f"Elia {ds} sample", f"https://opendata.elia.be/api/explore/v2.1/catalog/datasets/{ds}/records", limit=2)

show("Energinet GenerationProdTypeExchange", "https://api.energidataservice.dk/dataset/GenerationProdTypeExchange", n=1500,
     limit=3, start="2025-01-01", end="2025-01-02")
show("Energinet Gasflow", "https://api.energidataservice.dk/dataset/Gasflow", n=1200, limit=3)
show("Energinet GasSystemCommercialBalance", "https://api.energidataservice.dk/dataset/GasSystemCommercialBalance", n=1200, limit=3)
show("Energinet CapacityPerMunicipality", "https://api.energidataservice.dk/dataset/CapacityPerMunicipality", n=800, limit=2)
show("Energinet ProductionConsumptionSettlement last", "https://api.energidataservice.dk/dataset/ProductionConsumptionSettlement",
     n=300, limit=1, sort="HourUTC DESC")
show("NESO count", "https://api.neso.energy/api/3/action/datastore_search", n=300,
     resource_id="f93d1835-75bc-43e5-84ad-12472b180a98", limit=1, sort="DATETIME desc")
show("REE balance month DE detail", "https://apidatos.ree.es/en/datos/balance/balance-electrico", n=200,
     start_date="2024-01-01T00:00", end_date="2024-01-31T23:59", time_trunc="day", geo_trunc="electric_system", geo_limit="peninsular", geo_ids="8741")
