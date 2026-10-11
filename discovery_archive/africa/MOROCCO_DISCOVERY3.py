"""Morocco pass 3: follow links on one.org.ma, anre.ma, data.gov.ma CKAN; retry mem.gov.ma with relaxed TLS."""
import re, ssl, json, requests, urllib3
from requests.adapters import HTTPAdapter
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
class Legacy(HTTPAdapter):
    def init_poolmanager(self, *a, **k):
        c = ssl.create_default_context(); c.check_hostname = False; c.verify_mode = ssl.CERT_NONE
        c.set_ciphers("DEFAULT:@SECLEVEL=0"); c.options |= 0x4
        k["ssl_context"] = c; super().init_poolmanager(*a, **k)
S = requests.Session(); S.mount("https://", Legacy())
def get(u, **kw):
    try: return S.get(u, headers=H, timeout=(10, 25), **kw)
    except Exception as e: print("  ERR", u, type(e).__name__, str(e)[:120])
def links(u, pat=r".", n=80):
    print("\n=== LINKS", u)
    r = get(u)
    if r is None: return None
    print(" status", r.status_code, len(r.content), r.url)
    seen = set()
    for h, t in re.findall(r'href=["\']([^"\']+)["\'][^>]*>\s*([^<]{0,90})', r.text):
        if re.search(pat, h + " " + t, re.I) and (h, t) not in seen:
            seen.add((h, t)); print("   ", h[:150], "|", t.strip()[:70])
            if len(seen) >= n: break
    return r
links("https://www.one.org.ma/", n=80)
links("https://anre.ma/publications/", r"pdf|rapport|statist|bilan|annuel|donn|xls|production|march", 80)
links("https://anre.ma/", r"pdf|rapport|statist|bilan|donn|xls|open|chiffre|march", 60)
links("https://www.masen.ma/fr", r"open|donn|data|product|chiffre|statist|rapport", 50)
for u in ["https://www.mem.gov.ma/", "http://www.mem.gov.ma/", "https://www.mem.gov.ma/Pages/Statistiques.aspx"]:
    r = get(u); print(u, r.status_code if r is not None else None)
    if r is not None and r.status_code == 200:
        links(u, r"stat|bulletin|conjoncture|bilan|xls|pdf|donn", 60)
r = get("https://data.gov.ma/data/fr/api/3/action/package_search?q=energie&rows=20")
if r is not None and r.ok:
    for d in r.json()["result"]["results"]:
        print("DS", d["name"], d["metadata_modified"][:10], d["organization"]["title"] if d.get("organization") else "")
        for x in d["resources"]: print("    ", x["format"], x["url"][:130])
for q in ["onee", "electrique", "production electrique", "gaz", "petrole", "bilan energetique", "masen", "renouvelable"]:
    r = get(f"https://data.gov.ma/data/fr/api/3/action/package_search?q={q.replace(' ','+')}&rows=20")
    if r is not None and r.ok:
        j = r.json()["result"]; print("Q", q, j["count"], [d["name"] for d in j["results"]][:12])
for u in ["https://apidatos.ree.es/es/datos/intercambios/enlaces-internacionales?start_date=2026-09-01T00:00&end_date=2026-09-30T23:59&time_trunc=day",
          "https://apidatos.ree.es/en/datos/intercambios/todas-fronteras-fisicas?start_date=2026-09-01T00:00&end_date=2026-09-30T23:59&time_trunc=day",
          "https://apidatos.ree.es/en/datos/intercambios/marruecos?start_date=2026-09-01T00:00&end_date=2026-09-30T23:59&time_trunc=day",
          "https://api.energy-charts.info/public_power?country=ma&start=2026-09-01&end=2026-09-05",
          "https://api.energy-charts.info/cbpf?country=ma&start=2026-09-01&end=2026-09-05"]:
    r = get(u); print("\n", u[:110], r.status_code if r is not None else None)
    if r is not None: print(r.text[:600])
