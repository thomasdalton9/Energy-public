"""Probe 2: ARE Benin documents API params (manual). Compact output."""
import json, re, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "application/json, text/html, */*"}
API = "https://backoffice.are.bj/api/documents"
SLUG = "rapports-de-production-journalier-de-la-sbpe"

def get(url, **kw):
    try: return requests.get(url, headers=H, timeout=40, **kw)
    except Exception as e: print("ERR", url, e)

r = get(API)
j = r.json()
print("top keys", list(j.keys()), "n", len(j["data"]))
for k in j:
    if k != "data": print(k, json.dumps(j[k])[:700])
print("first item keys", list(j["data"][0].keys()))

# bundles: find api usage
page = get("https://www.are.bj/documents/categorie/" + SLUG).text
srcs = sorted(set(re.findall(r'/_next/static/chunks/[0-9a-f]+\.js', page)))
print(len(srcs), "chunks")
for s in srcs:
    t = get("https://www.are.bj" + s)
    if t is None: continue
    t = t.text
    for m in re.finditer(r'(api/documents|/documents|categorie|category_slug|per_page|backoffice)', t):
        a = max(0, m.start() - 150); print("JS", s[-14:], t[a:m.start()+250].replace("\n", " ")); 
        break
    for m in re.finditer(r'api/documents', t):
        a = max(0, m.start() - 300); print("JSDOC", s[-14:], t[a:m.start()+500].replace("\n", " "))
# RSC data in html around SBPE docs
for m in re.finditer(r'category_slug', page):
    print("HTMLCAT", page[max(0, m.start()-400):m.start()+200].replace("\\", ""))
    break

def trial(params):
    r = get(API, params=params)
    if r is None: return
    try:
        j = r.json(); d = j.get("data", [])
        cats = sorted({x.get("category_slug") for x in d})
        print("TRY", params, r.status_code, "n", len(d), "cats", cats[:4], "meta", json.dumps(j.get("meta"))[:200], "links", json.dumps(j.get("links"))[:200])
    except Exception as e:
        print("TRY", params, r.status_code, r.text[:200])

cands = []
for k in ["category", "category_slug", "categorie", "categorie_slug", "category_id", "categories", "slug", "type", "filter[category]", "filter[category_slug]", "filters[category][slug]", "cat"]:
    cands.append({k: SLUG}); cands.append({k: SLUG + "-documents"})
cands += [{"page": 2}, {"per_page": 100}, {"limit": 100}, {"perPage": 100}, {"page_size": 100}, {"search": "SBPE"}, {"q": "SBPE"}, {"query": "SBPE"}, {"keyword": "SBPE"}, {"s": "SBPE"}, {"title": "SBPE"}]
for c in cands: trial(c)
