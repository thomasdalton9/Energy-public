"""Probe 5."""
import io
import re
import zipfile
import requests
import pandas as pd

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u, **kw):
    try:
        return requests.get(u, headers=H, timeout=(10, 90), **kw)
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u} {type(e).__name__} {str(e)[:100]}", flush=True)


def txt(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def links(r, pat=r"."):
    return [(h, re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", l))[:70]) for h, l in
            re.findall(r'href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I) if re.search(pat, h, re.I)]




print("=== AESO asset csv tail")
u = "https://www.aeso.ca/assets/Uploads/Hourly-Metered-Volumes-by-Generating-Asset.csv"
r = requests.get(u, headers={**H, "Range": "bytes=306700000-"}, timeout=60)
print(r.status_code, r.headers.get("content-range"), r.content[-300:])
print("=== HQ hydrometeo")
base = "https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/donnees-hydrometeorologiques/records"
r = get(base, params={"limit": 6, "order_by": "date desc"})
print(r.status_code, r.text[:1800])
r = get(base, params={"limit": 8, "where": "nom like 'Réservoir' OR nom like 'Reservoir' OR nom like 'Caniapiscau'"})
print(r.status_code, r.text[:1500])
r = get("https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/donnees-hydrometriques/records", params={"limit": 3})
print(r.status_code, r.text[:1500])
print("=== BC reservoir raw")
r = get("https://www.bchydro.com/energy-in-bc/operations/transmission-reservoir-data/previous-reservoir-elevations/columbia.html")
t = r.text
for m in re.finditer(r"(<img[^>]+(?:chart|reservoir|elev)[^>]*>|<iframe[^>]*>|[^\"\']+\.(?:json|csv|png|gif)[^\"\']*)", t, re.I):
    print("   ", m.group(0)[:220])
i = t.find("Revelstoke")
print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t[t.find("Columbia Region", 20000):][:6000]))[:1500])
print("=== MB app")
for u in ["https://www.hydro.mb.ca/hydrologicalData/static/", "https://www.hydro.mb.ca/hydrologicalData/static/index.html"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, r.text[:500])
        print(re.findall(r'(?:src|href)="([^"]+)"', r.text)[:10])
print("=== SK scripts")
r = get("https://www.saskpower.com/about-us/our-company/power-system/system-data")
print(r.status_code, sorted(set(re.findall(r'src="([^"]+\.js[^"]*)"', r.text)))[:15])
print(txt(r.text)[-1500:])
r = get("https://www.saskpower.com/our-power-future/our-electricity-supply-and-demand")
print(txt(r.text)[2000:3500])
print("=== AESO CSD parse test")
r = get("http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet")
print(r.text[:1500])
