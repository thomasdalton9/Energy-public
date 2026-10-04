"""National Gas Data Portal: walk /api/find-gas-data-folders and search-everywhere for storage items (site-level stock / withdrawals / injections)."""
import json, re, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Referer": "https://data.nationalgas.com/find-gas-data/view", "Content-Type": "application/json"}
B = "https://data.nationalgas.com"
def show(tag, r, n=1500):
    print(tag, r.status_code, len(r.text)); print(r.text[:n].replace("\n", " "))
try:
    r = requests.get(B + "/api/find-gas-data-folders", headers=H, timeout=60); show("folders", r, 800)
    txt = r.text
    # print every object whose name mentions storage / rough, with its id
    for m in re.finditer(r'\{[^{}]*(?:[Ss]torage|[Rr]ough)[^{}]*\}', txt):
        print("HIT", m.group(0)[:300])
except Exception as e: print("folders ERR", e)
for q in ["storage", "Rough", "stock"]:
    for method in ("get", "post"):
        try:
            r = requests.get(B + "/api/search-everywhere", params={"query": q, "searchText": q, "q": q}, headers=H, timeout=60) if method == "get" else \
                requests.post(B + "/api/search-everywhere", json={"query": q, "searchText": q}, headers=H, timeout=60)
            show(f"search {q} {method}", r, 1200)
        except Exception as e: print("search ERR", e)
