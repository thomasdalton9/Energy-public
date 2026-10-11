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
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

def probe(u):
    out = []
    t0 = time.time()
    try:
        r = requests.get(u, headers=H, timeout=(8, 12), allow_redirects=True, stream=True)
        content = b""
        for chunk in r.iter_content(65536):
            content += chunk
            if len(content) > 400000 or time.time() - t0 > 25:
                break
        t = r.headers.get("content-type", "")
        txt = content.decode("utf-8", "ignore") if ("text" in t or "json" in t or "xml" in t) else ""
        title = re.search(r"<title[^>]*>(.*?)</title>", txt, re.S | re.I)
        out.append(f"{r.status_code} {len(content):>8} {t[:40]:40} {u} -> {r.url} | {title.group(1).strip()[:80] if title else ''}")
        if r.status_code == 200 and "text/html" in t:
            links = set(re.findall(r'href=["\']([^"\']+)["\']', txt))
            keep = [l for l in links if re.search(r"(?i)(dispatch|generat|statist|report|\.xls|\.pdf|dashboard|data|hydro|akosombo|lake|reservoir|outlook|load)", l)]
            for l in sorted(keep)[:60]:
                out.append("      link: " + l)
        if r.status_code == 200 and ("json" in t or "xml" in t):
            out.append("      body: " + txt[:600].replace("\n", " "))
    except Exception as e:
        out.append(f"ERR {u} {type(e).__name__}: {str(e)[:100]}")
    return "\n".join(out)

with ThreadPoolExecutor(8) as ex:
    for f in as_completed([ex.submit(probe, u) for u in URLS]):
        print(f.result(), flush=True)
