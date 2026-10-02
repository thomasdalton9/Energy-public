"""One-off probe: find EIA's US-Mexico border crossing capacity table. Logs only (nothing committed, no keys)."""
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (energy-data research)"}
PAGES = ["https://www.eia.gov/naturalgas/data.php", "https://www.eia.gov/naturalgas/pipelines/",
         "https://www.eia.gov/naturalgas/pipelines/EIA-NaturalGasPipelines.php",
         "https://www.eia.gov/pub/oil_gas/natural_gas/analysis_publications/ngpipeline/index.html",
         "https://www.eia.gov/naturalgas/weekly/"]
KEY = re.compile(r"border|mexic|capacity|crossing|xls|csv", re.I)
for u in PAGES:
    try:
        r = requests.get(u, headers=UA, timeout=60)
        print(f"## {u} -> {r.status_code} {len(r.content)}")
        links = sorted(set(re.findall(r'href="([^"]+)"[^>]*>([^<]{3,90})<', r.text)))
        for h, t in links:
            if KEY.search(h) or KEY.search(t):
                print("  ", t.strip()[:80], "|", h[:150])
    except Exception as e:
        print("##", u, "ERROR", type(e).__name__)
