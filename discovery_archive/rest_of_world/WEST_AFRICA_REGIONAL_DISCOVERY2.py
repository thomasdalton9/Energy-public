"""Round 2: drill into ecowapp.org, energydata.info (ECREEE/WAPP), AEP database, ecowrex, OMVG, ERERA."""
import re, json, requests
from urllib.parse import urljoin
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) Chrome/124 Safari/537.36"
def get(u, **k):
    try: return S.get(u, timeout=30, **k)
    except Exception as e: print("ERR", u, type(e).__name__); return None
def links(r, pat=r".", n=40):
    out = []
    for h, t in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.I | re.S):
        t = re.sub(r"<[^>]+>", "", t).strip()[:70]
        if re.search(pat, h + t, re.I) and (h, t) not in out: out.append((h, t))
    for h, t in out[:n]: print("   ", urljoin(r.url, h), "|", t)
print("=== ecowapp.org")
for u in ["https://www.ecowapp.org/", "https://www.ecowapp.org/en/publications", "https://www.ecowapp.org/en/information-and-coordination-center", "https://www.ecowapp.org/en/data", "https://www.ecowapp.org/en/statistics", "https://www.ecowapp.org/en/documents"]:
    r = get(u)
    if r is not None:
        print("##", u, r.status_code, r.url, len(r.content)); 
        if r.status_code == 200: links(r, r"ppm|bulletin|report|data|stat|dispatch|exchange|pdf|xls|csv|icc|coordination|dashboard|scada|live|operation", 45)
print("=== a PPM pdf (first 600 chars of text via header check)")
r = get("https://www.ecowapp.org/sites/default/files/wapp_ppm_ndeg1_de_lannee_2026.pdf")
if r is not None: print(r.status_code, r.headers.get("content-type"), len(r.content), r.content[:8])
print("=== energydata.info")
for q in ["ECREEE", "West African Power Pool", "WAPP", "ECOWAS electricity", "OMVS Manantali", "Volta River Authority", "electricity generation West Africa"]:
    r = get(f"https://energydata.info/api/3/action/package_search?q={requests.utils.quote(q)}&rows=15")
    if r is None or r.status_code != 200: continue
    j = r.json()["result"]; print("# q=", q, "count", j["count"])
    for p in j["results"]:
        fm = sorted({x.get("format", "") for x in p.get("resources", [])})
        print("  ", p["title"][:80], "|", p.get("end_date"), "|", p.get("license_id"), "|", fm, "|", "https://energydata.info/dataset/" + p["name"])
print("=== AEP")
for u in ["https://africa-energy-portal.org/database", "https://africa-energy-portal.org/utility-database", "https://africa-energy-portal.org/focus-area/generation", "https://africa-energy-portal.org/reports"]:
    r = get(u)
    if r is None: continue
    print("##", u, r.status_code)
    for m in sorted(set(re.findall(r'(?:src|href|action|data-[a-z-]+)="([^"]*(?:api|json|ajax|download|export|csv|xls|views|chart|highcharts)[^"]*)"', r.text, re.I)))[:30]: print("   attr:", m)
    for m in sorted(set(re.findall(r'drupalSettings"?[^<]{0,300}', r.text)))[:2]: print("   ds:", m[:300])
    for m in sorted(set(re.findall(r'https?://[^"\'\s<>]*\.(?:xlsx?|csv|json)', r.text)))[:15]: print("   file:", m)
for u in ["https://africa-energy-portal.org/sitemap.xml", "https://africa-energy-portal.org/robots.txt", "https://africa-energy-portal.org/jsonapi", "https://africa-energy-portal.org/_flysystem", "https://africa-energy-portal.org/data/export"]:
    r = get(u); print("##", u, r.status_code if r is not None else None, (r.text[:200].replace("\n", " ") if r is not None and r.status_code == 200 else ""))
print("=== ecowrex / ecreee data / pipes / omvg / erera / abn sath")
for u, pat in [("https://ecowrex.org/dashboard/", r"api|data|csv|xls|json|download|dashboard"), ("https://www.ecreee.org/software-research-3/data/", r"data|xls|csv|download|ecowrex|dashboard"),
               ("https://www.omvg.org/bulletins-dinformation", r"pdf|bulletin|energ|kaleta|sambangalou"), ("https://erera.arrec.org/", r"."), ("https://www.arrec.org/", r"data|report|exchange|xls|csv"),
               ("http://sath.abn.ne", r"api|json|csv|data|station"), ("https://www.omvs.org/?s=manantali", r"manantali|energie|production|pdf"), ("https://www.sogem-omvs.org/", r"."), ("https://www.sogem.sn/", r"."), ("https://www.vra.com/", r"data|report|generation|xls|csv"), ("https://www.ecowas.int/", r"energy|data")]:
    r = get(u)
    if r is None: continue
    print("##", u, r.status_code, r.url, len(r.content))
    if r.status_code == 200: links(r, pat, 25)
print("=== other ENTSO-E-style / bulk")
for u in ["https://ember-energy.org/", "https://api.worldbank.org/v2/country/NGA/indicator/EG.ELC.NGAS.ZS?format=json&per_page=3", "https://www.irena.org/-/media/Files/IRENA/Agency/Statistics/Statistical_Profiles/Africa/Nigeria_Africa_RE_SP.pdf",
          "https://www.eia.gov/opendata/", "https://api.eia.gov/v2/international/data/?frequency=annual&data[0]=value&facets[countryRegionId][]=NGA&facets[productId][]=2&length=2"]:
    r = get(u); print("##", u, r.status_code if r is not None else None, (r.headers.get("content-type", "")[:30] if r is not None else ""))
