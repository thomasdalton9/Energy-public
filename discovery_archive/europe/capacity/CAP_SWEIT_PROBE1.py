"""Probe 1: candidate monthly installed-capacity sources for FR, ES, PT, IT (run in Actions).
Findings are written in the docstring of CAP_SWEIT_PROBE2.py once known."""
import json, requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"}
URLS = [
 "https://apidatos.ree.es/en/datos/generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2026-09-30T23:59&time_trunc=month",
 "https://apidatos.ree.es/en/datos/generacion/potencia-instalada?start_date=2020-12-01T00:00&end_date=2021-03-31T23:59&time_trunc=month",
 "https://api.energy-charts.info/installed_power?country=fr&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=es&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=pt&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/installed_power?country=it&time_step=monthly&installation_decommission=false",
 "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=100&where=search(title,%22capacit%C3%A9%22)%20or%20search(title,%22puissance%20install%22)%20or%20search(title,%22registre%22)&select=dataset_id,metas",
 "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=100&where=search(dataset_id,%22registre%22)&select=dataset_id",
 "https://data.statistiques.developpement-durable.gouv.fr/dataset/?q=puissance+raccord",
 "https://www.statistiques.developpement-durable.gouv.fr/media/search?q=parc+installe",
 "https://datahub.ren.pt/en/electricity/installed-capacity/",
 "https://datahub.ren.pt/service/download/csv/?type=installedcapacity",
 "https://www.dgeg.gov.pt/pt/estatistica/energia/eletricidade/",
 "https://dgeg.gov.pt/pt/estatistica/energia/estatisticas-rapidas/",
 "https://dgeg.gov.pt/pt/estatistica/energia/eletricidade/",
 "https://www.terna.it/it/sistema-elettrico/statistiche/pubblicazioni-statistiche",
 "https://www.terna.it/en/electric-system/statistics/statistical-publications",
 "https://dati.terna.it/",
 "https://dati.terna.it/en/generation/installed-capacity",
 "https://www.terna.it/it/sistema-elettrico/dispacciamento/dati-esercizio",
 "https://developer.terna.it/",
 "https://download.terna.it/terna/",
 "https://www.gse.it/dati-e-scenari/statistiche",
 "https://www.gse.it/documenti_site/Documenti%20GSE/Rapporti%20statistici/",
 "https://www.entsoe.eu/",
]
import subprocess
for u in URLS:
    try:
        o = subprocess.run(["curl", "-sL", "-m", "30", "-A", UA["User-Agent"], "-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u],
                           capture_output=True, text=True, timeout=45)
        t = o.stdout
        print("\n=====", u, "| rc", o.returncode, "|", t[t.rfind("#HTTP"):].strip(), flush=True)
        print(t[:700].replace("\n", " "), flush=True)
    except Exception as e:
        print("\n=====", "ERR", type(e).__name__, u, flush=True)
