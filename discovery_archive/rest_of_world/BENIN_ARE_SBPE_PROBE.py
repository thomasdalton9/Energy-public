"""Probe ARE Benin document library for SBPE daily production reports (manual)."""
import json, re, sys, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
     "Accept": "application/json, text/html, */*"}
API = "https://backoffice.are.bj/api"
SLUG = "rapports-de-production-journalier-de-la-sbpe"

def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=40, **kw)
        return r
    except Exception as e:
        print("ERR", url, e); return None

def show(url, n=1500, **kw):
    r = get(url, **kw)
    if r is None: return None
    print("\n==", r.url, r.status_code, r.headers.get("content-type"), len(r.content))
    print(r.text[:n])
    return r

page = show("https://www.are.bj/documents/categorie/" + SLUG, 300)
if page is not None:
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', page.text, re.S)
    if m:
        print("NEXT_DATA len", len(m.group(1))); print(m.group(1)[:4000])
    else:
        print("no NEXT_DATA; scripts:")
        for s in re.findall(r'src="([^"]+\.js[^"]*)"', page.text)[:40]: print(s)
        # app router flight data
        for mm in re.findall(r'self\.__next_f\.push\(\[1,"(.{0,3000})', page.text)[:6]: print("FLIGHT", mm[:1500])
        for k in re.finditer(r'sbpe', page.text, re.I):
            print("CTX", page.text[max(0,k.start()-300):k.start()+300].replace("\n"," ")); break

r = show(API + "/documents", 2500)
for ep in ["categories", "document-categories", "categorie-documents", "categories-documents", "category-documents", "document-category", "categories?pagination[pageSize]=100"]:
    show(f"{API}/{ep}", 1200)
