"""Probe: do Elering, Conexus, Fluxys, Bulgartransgaz and Creos answer from GitHub, and which data/API links do their transparency
pages expose (for gas consumption and cross-border flows)? Prints status, size and links containing api/csv/xlsx/json/data."""
import re
import requests
H = {"User-Agent": "Mozilla/5.0", "Accept": "*/*"}
URLS = [
    "https://elering.ee/en/gas-balance", "https://gaas.elering.ee/", "https://dashboard.elering.ee/en/gas",
    "https://dashboard.elering.ee/api/gas/balance", "https://dashboard.elering.ee/swagger-ui/index.html", "https://dashboard.elering.ee/v3/api-docs",
    "https://www.conexus.lv/transparency", "https://transparency.conexus.lv/", "https://datu-parvaldiba.conexus.lv/",
    "https://www.gasdata.fluxys.com/", "https://gasdata.fluxys.com/", "https://www.fluxys.com/en/products-services/transmission/transparency",
    "https://www.bulgartransgaz.bg/en/pages/transparency", "https://www.bulgartransgaz.bg/en/pages/ntg-data", "https://www.bulgartransgaz.bg/en/pages/balancing-gas-market",
    "https://www.creos-net.lu/en/gas", "https://www.creos-net.lu/en/creos-luxembourg/open-data", "https://opendata.creos.lu/", "https://www.data.public.lu/en/search/?q=gaz",
]
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=25)
        links = sorted(set(re.findall(r'["\'(]([^"\'()\s<>]*(?:api|csv|xlsx|json|download|export|transparen|consumption|balance)[^"\'()\s<>]*)["\']?', r.text, re.I)))[:25]
        print(r.status_code, len(r.text), u, flush=True)
        for l in links:
            print("    ", l[:150])
    except Exception as e:  # noqa: BLE001
        print("ERR", u, type(e).__name__, str(e)[:80], flush=True)
