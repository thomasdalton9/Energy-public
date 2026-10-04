"""Find site-level UK gas storage data (Rough, Hornsea, Aldbrough, Holford, Hill Top, Humbly Grove, Stublach, Hatfield Moor):
(1) National Gas Data Portal catalogue: scrape the portal JS for API endpoints and storage-related item names;
(2) GIE AGSI+ unit-level data for GB; (3) print what answers."""
import re, json, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36", "Accept": "*/*",
     "Referer": "https://data.nationalgas.com/find-gas-data/view"}
base = "https://data.nationalgas.com"
try:
    html = requests.get(base + "/find-gas-data", headers=H, timeout=60).text
    print("html", len(html))
    js = sorted(set(re.findall(r'src="([^"]+\.js)"', html)))
    print("js", js[:10])
    for s in js[:8]:
        u = s if s.startswith("http") else base + "/" + s.lstrip("/")
        t = requests.get(u, headers=H, timeout=60).text
        print(u, len(t))
        for m in sorted(set(re.findall(r'["\'](/?api/[A-Za-z0-9_\-/]+)["\']', t)))[:40]:
            print("   api:", m)
        for m in sorted(set(re.findall(r'[A-Za-z ,]*(?:Storage|Rough)[A-Za-z ,]*', t)))[:30]:
            if 6 < len(m) < 80: print("   name:", m.strip())
except Exception as e:
    print("portal ERR", e)
for path in ["/api/find-gas-data/catalogue", "/api/find-gas-data/items", "/api/find-gas-data/categories", "/api/data-items", "/api/find-gas-data/search?query=storage"]:
    try:
        r = requests.get(base + path, headers=H, timeout=60)
        print(path, r.status_code, r.text[:400].replace("\n", " "))
    except Exception as e:
        print(path, "ERR", e)
for q in ["storage stock", "rough storage"]:
    try:
        r = requests.post(base + "/api/find-gas-data", json={"latestFlag": "Y", "applicableFor": "Y", "dateFrom": "2025-01-01", "dateTo": "2025-01-02", "dateType": "GASDAY", "ids": "PUBOB4423"}, headers={**H, "Content-Type": "application/json"}, timeout=60)
        print("sample", r.status_code, r.text[:300])
    except Exception as e:
        print("sample ERR", e)
# AGSI+ GB (public endpoint needs key in header; try without)
try:
    r = requests.get("https://agsi.gie.eu/api", params={"country": "GB", "from": "2020-06-01", "to": "2020-06-03", "size": 5}, headers=H, timeout=60)
    print("agsi GB", r.status_code, r.text[:500])
except Exception as e:
    print("agsi ERR", e)
