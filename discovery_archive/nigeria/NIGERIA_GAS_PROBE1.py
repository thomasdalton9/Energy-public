"""Discovery: Nigerian gas sources reachability and structure (run in Actions)."""
import re, sys
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
URLS = [
 "https://www.nuprc.gov.ng/",
 "https://www.nuprc.gov.ng/oil-and-gas-production-data/",
 "https://www.nuprc.gov.ng/annual-statistical-bulletin/",
 "https://www.nuprc.gov.ng/monthly-oil-and-gas-production/",
 "https://nuprc.gov.ng/",
 "https://commercial.nuprc.gov.ng/",
 "https://www.nnpcgroup.com/",
 "https://nnpcgroup.com/Public-Relations/Monthly-Financial-Operations-Report",
 "https://www.nnpcgroup.com/Public-Relations/Monthly-Financial-Operations-Report",
 "https://www.nnpcgroup.com/operations/monthly-reports",
 "https://www.nigerialng.com/",
 "https://www.nigerialng.com/Pages/Facts-and-Figures.aspx",
 "https://www.nigerialng.com/investors",
 "https://nmdpra.gov.ng/",
 "https://www.nmdpra.gov.ng/",
 "https://www.nmdpra.gov.ng/domestic-base-price-dbp/",
 "https://nigerianstat.gov.ng/",
 "https://www.nigerianstat.gov.ng/elibrary",
 "https://nigerianstat.gov.ng/elibrary?page=1",
 "https://www.jodidata.org/gas/",
 "https://www.jodidata.org/gas/database/data-downloads.aspx",
 "https://www.jodidata.org/_resources/files/downloads/gas-data/gas_world_NewFormat.csv",
 "https://www.jodidata.org/_resources/files/downloads/gas-data/gas_world_csv.zip",
 "https://www.jodidata.org/_resources/files/downloads/gas-data/gas_primary_csv.zip",
 "https://www.nigeriagasflaretracker.org/",
 "https://www.ngfcp.gov.ng/",
 "https://www.nerc.gov.ng/",
 "https://nerc.gov.ng/index.php/library/documents/Quarterly-Reports/",
 "https://www.gasnigeria.net/",
 "https://nnpcgroup.com/Public-Relations/Monthly_Financial_Operations_Report",
]
for u in URLS:
    try:
        r = requests.get(u, headers=UA, timeout=25, allow_redirects=True)
        ct = r.headers.get("content-type", "")
        print(f"\n== {u}\n   {r.status_code} {ct} {len(r.content)}B final={r.url}")
        if "html" in ct or "text" in ct:
            t = r.text
            m = re.search(r"<title>(.*?)</title>", t, re.S | re.I)
            print("   title:", (m.group(1).strip()[:100] if m else None))
            links = set(re.findall(r'href=["\']([^"\']+\.(?:xlsx?|pdf|csv|zip)[^"\']*)', t, re.I))
            for l in sorted(links)[:40]:
                print("   file:", l)
            gl = set(re.findall(r'href=["\']([^"\']*(?:gas|production|report|statist|bulletin|flar)[^"\']*)', t, re.I))
            for l in sorted(gl)[:25]:
                print("   link:", l)
        else:
            print("   head:", r.content[:200])
    except Exception as e:
        print(f"\n== {u}\n   ERR {type(e).__name__}: {str(e)[:120]}")
    sys.stdout.flush()
