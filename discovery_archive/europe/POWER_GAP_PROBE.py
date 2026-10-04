"""Probe national power-generation sources (DE/IT/PL/RO/BG) reachable from GitHub Actions: status, size, head of body."""
import requests, json, sys
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
U = [
 "https://api.energy-charts.info/public_power?country=de&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/total_power?country=de&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/installed_power?country=de&time_step=monthly&installation_decommission=false",
 "https://api.energy-charts.info/public_power?country=it&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/public_power?country=pl&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/public_power?country=ro&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/public_power?country=bg&start=2024-06-01&end=2024-06-02",
 "https://api.energy-charts.info/ember_monthly?country=de",
 "https://www.smard.de/app/chart_data/1223/DE/index_month.json",
 "https://www.smard.de/app/chart_data/1223/DE/1223_DE_month_1704067200000.json",
 "https://api.raporty.pse.pl/api/his-wlk-cal?$filter=doba%20eq%20'2024-06-02'&$top=3",
 "https://api.raporty.pse.pl/api/his-gen-pal?$filter=doba%20eq%20'2024-06-02'&$top=3",
 "https://api.raporty.pse.pl/api/kse-load?$filter=doba%20eq%20'2024-06-02'&$top=3",
 "https://api.raporty.pse.pl/api/gen-jw?$filter=doba%20eq%20'2024-06-02'&$top=3",
 "https://api.raporty.pse.pl/api/his-gen-pal?$top=1",
 "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem?format=JSON&lang=EN&geo=IT&time=2024-06&siec=TOTAL&unit=GWH",
 "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem?format=JSON&lang=EN&geo=IT&time=2024-06",
 "https://api.terna.it/transparency/v1.0/getactualgeneration",
 "https://www.terna.it/it/sistema-elettrico/transparency-report/actual-generation",
 "https://dati.terna.it/",
 "https://download.terna.it/terna/",
 "https://www.mercatoelettrico.org/en/",
 "https://www.transelectrica.ro/en/web/tel/sistemul-energetic-national",
 "https://newmarkets.transelectrica.ro/uu-webkit-maga/00121002/00121002.pdf",
 "https://www.eso.bg/en/",
 "https://api.energidataservice.dk/dataset/ProductionConsumptionSettlement?limit=1",
 "https://energy-charts.info/charts/power/data/de/week_2024_23.json",
]
for u in U:
    try:
        r = requests.get(u, headers=H, timeout=40)
        print(r.status_code, len(r.content), u, flush=True)
        print("   ", r.text[:300].replace("\n", " "), flush=True)
    except Exception as e:
        print("ERR", u, type(e).__name__, str(e)[:100], flush=True)
