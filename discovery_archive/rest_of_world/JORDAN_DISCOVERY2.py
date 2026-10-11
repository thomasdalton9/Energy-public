"""Jordan power/gas discovery 2: probe NEPCO, EMRC, MEMR, DOS, open data portals, JODI, Aqaba."""
import re, sys, requests
requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

def get(url, show=400, links=True):
    print("\n=== " + url, flush=True)
    try:
        r = requests.get(url, headers=H, timeout=(10, 30), verify=False, allow_redirects=True)
    except Exception as e:
        print("  ERR", type(e).__name__, str(e)[:150]); return None
    print(f"  {r.status_code} {len(r.content)}B {r.headers.get('content-type')} final={r.url}")
    t = r.text if 'text' in (r.headers.get('content-type') or '') or 'json' in (r.headers.get('content-type') or '') else ''
    if t:
        print("  HEAD:", re.sub(r'\s+', ' ', t[:show]))
        if links:
            ls = sorted(set(re.findall(r'href=["\']([^"\']+)["\']', t)))
            keep = [l for l in ls if re.search(r'\.(xlsx?|csv|pdf|json)|api|download|resource|stat|report|load|generat|energy|electric|gas|bulletin|data', l, re.I)]
            for l in keep[:60]: print("   L", l)
    return r

urls = [
 "https://www.nepco.com.jo/", "https://www.nepco.com.jo/en/", "https://www.nepco.com.jo/en/index_en.aspx",
 "https://www.nepco.com.jo/en/annual_reports_en.aspx", "https://www.nepco.com.jo/en/electrical_energy_en.aspx",
 "https://www.nepco.com.jo/en/dailyloads_en.aspx", "https://www.nepco.com.jo/en/load_en.aspx",
 "https://www.nepco.com.jo/en/statistics_en.aspx", "https://www.nepco.com.jo/ar/index_ar.aspx",
 "https://www.emrc.gov.jo/", "https://www.emrc.gov.jo/en/", "https://emrc.gov.jo/en/statistics",
 "https://www.memr.gov.jo/", "https://www.memr.gov.jo/EN/Pages/Energy_Balance", "https://www.memr.gov.jo/EN/List/Annual_Reports",
 "https://www.memr.gov.jo/EN/List/Energy_Statistics",
 "https://dosweb.dos.gov.jo/", "https://jorinfo.dos.gov.jo/", "https://dosweb.dos.gov.jo/economic/energy/",
 "https://opendata.gov.jo/en/dataset/nepco-2170-2023", "https://opendata.gov.jo/api/3/action/package_search?q=electricity&rows=20",
 "https://opendata.gov.jo/api/3/action/package_search?q=nepco&rows=20",
 "https://data.gov.jo/", "https://www.jordanopendata.gov.jo/",
 "https://www.jodidata.org/gas/database/data-downloads.aspx",
 "https://www.nepco.com.jo/store/DOCS/web/2023_en.pdf",
 "https://www.jordantimes.com/news/local/nepco",
 "https://www.aqabaports.com/", "https://www.adc.jo/", "https://www.npc.com.jo/", "https://www.nrc.jo/",
]
for u in urls: get(u)
