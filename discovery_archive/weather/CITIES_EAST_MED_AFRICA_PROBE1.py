"""Probe 1 (Oct 2026): which official city-population sources are reachable from Actions for the East Med and Africa weather regions."""
import io
import os
import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "weather")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URLS = [
    "https://population.un.org/wup/Download/Files/WUP2018-F12-Cities_Over_300K.xls",
    "https://population.un.org/wup/assets/Download/WUP2018-F12-Cities_Over_300K.xls",
    "https://population.un.org/wup/Download/",
    "https://data.un.org/Data.aspx?d=POP&f=tableCode%3A240",
    "https://www.cbs.gov.il/en/publications/Pages/2024/Population-of-Localities.aspx",
    "https://data.gov.il/api/3/action/package_search?q=%D7%9B%D7%9C%D7%9B%D7%9C%D7%94+%D7%90%D7%95%D7%9B%D7%9C%D7%95%D7%A1%D7%99%D7%99%D7%94+%D7%99%D7%A9%D7%95%D7%91%D7%99%D7%9D&rows=10",
    "https://data.gov.il/api/3/action/package_search?q=localities+population&rows=10",
    "https://www.cbs.gov.il/he/publications/doclib/2024/yishuvim_2023/",
    "https://www.capmas.gov.eg/Pages/ShowPDF.aspx?page_id=/Admin/Pages%20Files/census2017.pdf",
    "https://www.capmas.gov.eg/",
    "https://censusinfo.capmas.gov.eg/",
    "https://www.cystat.gov.cy/en/",
    "https://www.cystat.gov.cy/en/SubthemeStatistics?id=1",
    "https://dosweb.dos.gov.jo/",
    "https://dosweb.dos.gov.jo/population/population-2/",
    "https://dosweb.dos.gov.jo/DataBank/Population/Population_Estimares/PopulationEstimatesofJordanbyLocality2023.pdf",
    "http://www.cas.gov.lb/",
    "http://www.cas.gov.lb/index.php/demographic-and-social-en/population-en",
    "https://www.ons.dz/",
    "https://www.hcp.ma/",
    "https://www.hcp.ma/downloads/RGPH-2024_t24425.html",
    "https://www.statssa.gov.za/",
    "https://www.statsghana.gov.gh/",
    "https://www.nigerianstat.gov.ng/",
    "https://ins.cm/",
    "https://api.worldbank.org/v2/country/EGY/indicator/SP.POP.TOTL?format=json&per_page=3",
    "https://download.geonames.org/export/dump/cities15000.zip",
]
lines = []
for u in URLS:
    try:
        r = requests.get(u, headers=UA, timeout=60)
        lines.append(f"{r.status_code}\t{len(r.content)}\t{r.headers.get('content-type','')}\t{u}\n    {r.text[:200].strip()!r}" if r.status_code == 200 and 'html' in r.headers.get('content-type','') or r.status_code != 200 else f"{r.status_code}\t{len(r.content)}\t{r.headers.get('content-type','')}\t{u}")
        if u.endswith("Cities_Over_300K.xls") and r.status_code == 200:
            open(os.path.join(OUT, "wup2018_f12.xls"), "wb").write(r.content)
    except Exception as e:  # noqa: BLE001
        lines.append(f"ERR\t{u}\t{e!r}"[:300])
    print(lines[-1], flush=True)
open(os.path.join(OUT, "cities_probe1.txt"), "w").write("\n".join(lines))
