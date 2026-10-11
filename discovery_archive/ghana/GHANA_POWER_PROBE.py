"""Probe Ghanaian power-generation sources (GRIDCo, Energy Commission, GSS, PURC, VRA, WAPP). Prints status, type, size, snippet."""
import re, sys, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URLS = [
 "https://www.gridcogh.com/", "https://www.gridcogh.com/dispatch-reports/", "https://gridcogh.com/", "https://www.gridcogh.com/reports/",
 "https://www.gridcogh.com/operations/", "https://www.gridcogh.com/system-operations/", "https://www.gridcogh.com/wp-json/wp/v2/pages?per_page=100",
 "https://www.gridcogh.com/wp-json/wp/v2/media?per_page=100&search=generation",
 "https://www.gridcogh.com/wp-sitemap.xml", "https://www.gridcogh.com/sitemap.xml",
 "https://www.energycom.gov.gh/", "https://www.energycom.gov.gh/planning/data-center/energy-statistics",
 "https://energycom.gov.gh/", "https://www.energycom.gov.gh/files/", "https://energycom.gov.gh/planning/data-center",
 "https://www.statsghana.gov.gh/", "https://statsghana.gov.gh/", "https://www.purc.com.gh/", "https://purc.com.gh/",
 "https://www.vra.com/", "https://vra.com/", "https://www.vra.com/resources/", "https://www.vra.com/our_mandate/hydro.php",
 "https://www.ecowapp.org/", "https://www.ecowapp.org/en/", "https://www.ecowapp.org/en/documents", "https://www.ecowapp.org/en/data",
 "https://www.nitag.gov.gh/", "https://www.mope.gov.gh/", "https://energymin.gov.gh/",
 "https://www.gridcogh.com/real-time-data/", "https://www.gridcogh.com/generation/", "https://www.gridcogh.com/data/",
 "https://dashboard.gridcogh.com/", "https://www.gridcogh.com/system-dashboard/",
 "https://www.energycom.gov.gh/planning/data-center/energy-outlook", "https://www.energycom.gov.gh/planning/data-center/national-energy-statistics",
 "https://data.gov.gh/", "https://www.ember-energy.org/",  # skip ember data, just reachability
 "https://www.wapp-ecowas.org/", "https://www.ecowapp.org/en/publications",
 "https://www.irena.org/", "https://opendata.gridcogh.com/",
]
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=25, allow_redirects=True)
        t = r.headers.get("content-type", "")
        txt = r.text[:3000] if "text" in t or "json" in t or "xml" in t else ""
        title = re.search(r"<title[^>]*>(.*?)</title>", txt, re.S | re.I)
        print(f"{r.status_code} {len(r.content):>8} {t[:40]:40} {u} -> {r.url} | {title.group(1).strip()[:80] if title else ''}", flush=True)
        if r.status_code == 200 and "text/html" in t:
            full = r.text
            links = set(re.findall(r'href=["\']([^"\']+)["\']', full))
            keep = [l for l in links if re.search(r"(?i)(dispatch|generat|statist|report|\.xls|\.pdf|dashboard|data|hydro|akosombo|lake|reservoir|outlook|load)", l)]
            for l in sorted(keep)[:60]:
                print("      link:", l, flush=True)
        if r.status_code == 200 and ("json" in t or "xml" in t):
            print("      body:", r.text[:600].replace("\n", " "), flush=True)
    except Exception as e:
        print(f"ERR {u} {type(e).__name__}: {str(e)[:100]}", flush=True)
