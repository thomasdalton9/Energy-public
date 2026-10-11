"""Discovery: regional West African grid data (WAPP, ERERA, ECREEE, OMVS/OMVG, VBA/NBA, Africa Energy Portal)."""
import re, json, sys, requests
from urllib.parse import urljoin
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) Chrome/124 Safari/537.36"
URLS = [
 "https://www.wapp-ecowas.org/", "https://www.wapp-ecowas.org/index.php/en/", "https://www.wapp-ecowas.org/publications",
 "https://www.wapp-ecowas.org/wp-json/wp/v2/pages?per_page=5",
 "https://www.ecowapp.org/", "https://www.erera.arrec.org/", "https://erera.arrec.org/", "https://www.erera.org/",
 "https://www.ecreee.org/", "https://www.ecreee.org/data/", "https://www.ecowrex.org/", "https://www.ecowrex.org/dataset",
 "https://www.omvs.org/", "https://www.omvg.org/", "https://www.omvs.org/energie", "https://www.sogem.org/",
 "https://www.voltabasin.org/", "https://www.vba-oibv.org/", "https://www.abn.ne/", "https://www.abn.ne/en",
 "https://africa-energy-portal.org/", "https://africa-energy-portal.org/database", "https://africa-energy-portal.org/api",
 "https://africa-energy-portal.org/api/data", "https://africa-energy-portal.org/sites/default/files",
 "https://africa-energy-portal.org/aep/country/nigeria", "https://africa-energy-portal.org/aep/energy-sources/generation",
 "https://www.africa-energy-portal.org/swagger", "https://api.africa-energy-portal.org/",
 "https://transparency.entsoe.eu/", "https://data.worldbank.org/", 
 "https://www.irena.org/Data/Download-Data", "https://pxweb.irena.org/pxweb/en/IRENASTAT",
 "https://api.worldbank.org/v2/country/NGA;GHA;SEN;CIV/indicator/EG.ELC.PROD.KH?format=json&per_page=5",
 "https://energydata.info/api/3/action/package_search?q=west+africa+power+pool&rows=10",
 "https://energydata.info/api/3/action/package_search?q=WAPP&rows=10",
 "https://data.humdata.org/api/3/action/package_search?q=west+africa+electricity&rows=10",
]
KEY = re.compile(r"(xlsx?|csv|json|api|download|dispatch|exchange|flow|generation|statistic|data|report|bulletin|manantali|felou|kaleta|akosombo)", re.I)
def probe(u):
    try:
        r = S.get(u, timeout=25, allow_redirects=True)
    except Exception as e:
        print(f"## {u}\n   ERR {type(e).__name__}: {str(e)[:120]}"); return
    ct = r.headers.get("content-type", "")
    print(f"## {u}\n   {r.status_code} {ct[:50]} {len(r.content)}B final={r.url}")
    if r.status_code != 200: return
    t = r.text
    if "json" in ct:
        print("   JSON:", t[:600].replace("\n", " ")); return
    m = re.search(r"<title[^>]*>(.*?)</title>", t, re.I | re.S)
    print("   title:", (m.group(1).strip()[:100] if m else None))
    seen = set()
    for h, txt in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', t, re.I | re.S):
        txt = re.sub(r"<[^>]+>", "", txt).strip()[:70]
        if (KEY.search(h) or KEY.search(txt)) and h not in seen and len(seen) < 25:
            seen.add(h); print("   link:", urljoin(r.url, h), "|", txt)
for u in URLS: probe(u)
