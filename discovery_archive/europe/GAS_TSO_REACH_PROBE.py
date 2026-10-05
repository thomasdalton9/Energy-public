"""
Probe: which gas TSO / statistics sites answer from GitHub Actions (status, content type, size, first bytes), to cross-check ENTSOG border flows
against operator data: Eustream, FGSZ, MEKH, Gaz-System, Transgaz, Bulgartransgaz, Plinacro, Plinovodi, Gasgrid, Conexus, Amber Grid, GTSOU.
"""
import requests

URLS = [
    "https://www.eustream.sk/en/transparency/flow-data", "https://www.eustream.sk/en/transparency/data-and-publications",
    "https://data.eustream.sk/", "https://transparency.eustream.sk/en/",
    "https://mgp.fgsz.hu/en", "https://www.fgsz.hu/en/transparency", "https://ugyfel.fgsz.hu/",
    "https://mekh.hu/en/gas-statistics", "https://www.mekh.hu/download/",
    "https://www.gaz-system.pl/en/transmission-system/transmission-services/transparency/", "https://swo.gaz-system.pl/",
    "https://www.transgaz.ro/en/transparency", "https://www.transgaz.ro/en/",
    "https://www.bulgartransgaz.bg/en/pages/transparency-1", "https://www.bulgartransgaz.bg/en/",
    "https://transparency.plinacro.hr/", "https://www.plinacro.hr/default.aspx?id=1049",
    "https://www.plinovodi.si/en/", "https://www.plinovodi.si/en/transparency/",
    "https://gasgrid.fi/en/", "https://gasgrid.fi/en/gas-market/transparency/",
    "https://transparency.conexus.lv/", "https://www.conexus.lv/",
    "https://transparency.ambergrid.lt/", "https://www.ambergrid.lt/en",
    "https://tsoua.com/en/", "https://tsoua.com/en/transparency/",
    "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_ti_gasm?format=JSON&lang=EN&geo=BG&partner=TR&time=2025-01&unit=TJ_GCV&siec=G3000",
]
H = {"User-Agent": "Mozilla/5.0"}
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=20)
        print(r.status_code, len(r.content), r.headers.get("content-type", "")[:40], u, "|", r.text[:100].replace("\n", " ") if r.status_code == 200 else "", flush=True)
    except Exception as e:  # noqa: BLE001
        print("ERR", type(e).__name__, u, flush=True)
