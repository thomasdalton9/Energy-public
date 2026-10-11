"""Round 3: AEP database backend endpoint, ecowapp.org site map, VRA facts."""
import re, requests
from urllib.parse import urljoin
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) Chrome/124 Safari/537.36"
def get(u, **k):
    try: return S.get(u, timeout=30, **k)
    except Exception as e: print("ERR", u, type(e).__name__); return None
print("=== ecowapp.org")
r = get("https://www.ecowapp.org/")
if r is not None:
    for h, t in sorted(set((h, re.sub(r"<[^>]+>", "", t).strip()[:50]) for h, t in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.I | re.S)))[:90]:
        print("  ", urljoin(r.url, h), "|", t)
print("=== AEP database page: inline JS / module js")
r = get("https://africa-energy-portal.org/database")
if r is not None:
    for m in re.findall(r"<script[^>]*src=\"([^\"]*aepdatabase[^\"]*)\"", r.text): print("script", m)
    for m in re.findall(r'(?:fetch|ajax|url)\s*[:(]\s*["\']([^"\']+)["\']', r.text)[:20]: print("js url", m)
    i = r.text.find("drupal-settings-json"); print(r.text[i:i+1500].replace("\n", " "))
    for js in set(re.findall(r'src="(/modules/custom/aepdatabase_mongo/[^"]*\.js[^"]*)"', r.text)):
        if "libraries" in js: continue
        j = get(urljoin(r.url, js)); print("## js", js, j.status_code if j is not None else None)
        if j is not None and j.status_code == 200:
            for m in sorted(set(re.findall(r'["\'](/[a-zA-Z0-9_/\-\.?=&{}]*(?:api|data|json|chart|get|download)[a-zA-Z0-9_/\-\.?=&{}]*)["\']', j.text)))[:30]: print("   ep", m)
print("=== AEP guess endpoints")
for u in ["https://africa-energy-portal.org/aepdatabase", "https://africa-energy-portal.org/aep/data", "https://africa-energy-portal.org/aepdatabase/get-data", "https://africa-energy-portal.org/database/data", "https://africa-energy-portal.org/aep/api/countries"]:
    x = get(u); print(u, x.status_code if x is not None else None, (x.headers.get("content-type", "")[:30], x.text[:150].replace("\n", " ")) if x is not None and x.status_code == 200 else "")
print("=== VRA facts")
r = get("https://vra.com/resources/facts.php")
if r is not None:
    print(r.status_code, len(r.content))
    t = re.sub(r"<[^>]+>", " ", r.text); t = re.sub(r"\s+", " ", t); print(t[:1500])
