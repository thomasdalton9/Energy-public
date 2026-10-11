"""Morocco power/gas discovery pass 2: broad URL probe (status, type, links)."""
import re, sys, requests
requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "fr,en;q=0.8"}
URLS = [
 "https://www.one.org.ma/", "https://www.one.org.ma/FR/", "https://www.onee.ma/", "https://www.onee.ma/fr",
 "https://www.onee.ma/fr/electricite/chiffres-cles", "https://www.onee.ma/fr/rapports-annuels",
 "https://www.onee.ma/fr/content/dispatching-national", "https://www.onee.ma/fr/content/open-data",
 "https://www.one.org.ma/FR/pages/interne.asp?esp=2&id1=3&id2=57&t2=1",
 "https://www.mem.gov.ma/", "https://www.mem.gov.ma/Pages/Secteur.aspx?e=2",
 "https://www.transitionenergetique.gov.ma/", "https://www.transitionenergetique.gov.ma/fr/statistiques",
 "https://www.energie.gov.ma/", "https://www.mem.gov.ma/Pages/ListeBulletin.aspx",
 "https://www.masen.ma/fr", "https://www.masen.ma/en/open-data",
 "https://anre.ma/", "https://www.anre.ma/", "https://www.anre.ma/fr/publications",
 "https://www.hcp.ma/", "https://www.hcp.ma/Indice-de-la-production-energetique_r98.html",
 "https://www.hcp.ma/Industrie-electrique_a2036.html",
 "https://www.data.gov.ma/", "https://data.gov.ma/data/fr/api/3/action/package_search?q=electricite",
 "https://data.gov.ma/data/fr/api/3/action/package_search?q=energie",
 "https://data.gov.ma/data/fr/dataset?q=onee",
 "https://www.bkam.ma/", "https://www.finances.gov.ma/fr/Pages/stat-bulletin.aspx",
 "https://www.finances.gov.ma/Publication/depf/",
 "https://www.onhym.com/", "https://www.soundenergy.ca/",
 "https://www.gme.ma/", "https://www.ree.es/en/datos/intercambios",
 "https://apidatos.ree.es/en/datos/intercambios/exportacion-importacion-energia?start_date=2026-09-01T00:00&end_date=2026-09-30T23:59&time_trunc=day",
 "https://www.iea.org/countries/morocco",
 "https://www.energy-charts.info/", "https://dashboard.oicp.ma/", "https://www.mem.gov.ma/Pages/Statistiques.aspx",
 "https://www.mem.gov.ma/Pages/Actualite.aspx", "https://www.mem.gov.ma/en/Pages/default.aspx",
]
def go(u):
    print("\n=== " + u)
    try:
        r = requests.get(u, headers=H, timeout=(10, 30), verify=False, allow_redirects=True)
    except Exception as e:
        print("  ERR", type(e).__name__, str(e)[:150]); return
    print("  status", r.status_code, "bytes", len(r.content), r.headers.get("content-type"), "final", r.url)
    if r.status_code == 200 and "json" in (r.headers.get("content-type") or ""):
        print(r.text[:1500]); return
    if r.status_code == 200:
        t = r.text
        m = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        print("  title:", (m.group(1).strip()[:120] if m else None))
        links = set()
        for h, txt in re.findall(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,80})<', t):
            s = (h + " " + txt).lower()
            if re.search(r"xls|csv|json|pdf|statisti|chiffre|bulletin|conjoncture|dispatch|open.?data|product|energie|electri|gaz|gas|import", s):
                links.add((h[:140], txt.strip()[:60]))
        for l in sorted(links)[:40]: print("   ", l)
for u in URLS: go(u)
