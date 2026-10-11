"""Round 4: AEP /get-database-data endpoint, WAPP ICC site, WAPP documentation list."""
import re, requests
from urllib.parse import urljoin
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) Chrome/124 Safari/537.36"
def get(u, **k):
    try: return S.get(u, timeout=30, **k)
    except Exception as e: print("ERR", u, type(e).__name__, str(e)[:80]); return None
print("=== AEP widget JS around get-database-data")
j = get("https://africa-energy-portal.org/modules/custom/aepdatabase_mongo/js/database-widget.js?v=1.0.0")
if j is not None:
    for m in re.finditer(r"get-database-data|get-homemap", j.text):
        print(j.text[max(0, m.start()-500):m.end()+600].replace("\n", " ")); print("-----")
        break
for u in ["https://africa-energy-portal.org/get-database-data", "https://africa-energy-portal.org/get-homemap-filter"]:
    r = get(u); print(u, r.status_code if r is not None else None, (r.headers.get("content-type", ""), r.text[:500].replace("\n", " ")) if r is not None else "")
print("=== ICC")
for u in ["http://icc.ecowapp.org", "https://icc.ecowapp.org", "http://pipes.ecowapp.org"]:
    r = get(u, allow_redirects=True)
    if r is None: continue
    print("##", u, r.status_code, r.url, len(r.content))
    if r.status_code == 200:
        m = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.S | re.I); print("title", m.group(1).strip()[:100] if m else None)
        for h, t in sorted(set((h, re.sub(r"<[^>]+>", "", t).strip()[:50]) for h, t in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.I | re.S)))[:40]: print("   ", urljoin(r.url, h), "|", t)
print("=== WAPP documentation")
r = get("https://www.ecowapp.org/en/documentation")
if r is not None:
    print(r.status_code)
    for h, t in sorted(set((h, re.sub(r"<[^>]+>", "", t).strip()[:60]) for h, t in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.I | re.S)))[:80]:
        if re.search(r"pdf|xls|csv|document|ppm|report|annual", h + t, re.I): print("   ", urljoin(r.url, h), "|", t)
