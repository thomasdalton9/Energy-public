"""National Gas Data Portal: list every catalogue item whose name/description mentions storage or Rough (id + name), then query each for a
few 2025 days and print the item names returned (site-level items show as separate itemName rows). Compact output."""
import re, json, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Referer": "https://data.nationalgas.com/find-gas-data/view", "Content-Type": "application/json"}
B = "https://data.nationalgas.com"
txt = requests.get(B + "/api/find-gas-data-folders", headers=H, timeout=60).text
items = {}
for m in re.finditer(r'"name":\s*"([^"]*)",\s*"description":\s*"((?:PUBOB?J?\d+)[^"]*)"', txt):
    name, desc = m.group(1), m.group(2)
    pid = re.match(r"(PUBOB?J?\d+)", desc).group(1)
    if re.search(r"storage|rough|stock|aldbrough|hornsea|holford|humbly|hill top|stublach", name + " " + desc, re.I):
        items[pid] = name
print(len(items), "storage-related catalogue items")
for k, v in items.items(): print(" ", k, "|", v)
ids = list(items)
for i in range(0, len(ids), 20):
    body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": "2025-01-10", "dateTo": "2025-01-11", "dateType": "GASDAY", "ids": ",".join(ids[i:i+20])}
    r = requests.post(B + "/api/find-gas-data", json=body, headers=H, timeout=120)
    print("query", r.status_code)
    seen = {}
    for it in r.json().get("data", []):
        seen.setdefault(it.get("itemName"), []).append((it.get("applicableFor"), it.get("value"), it.get("unitOfMeasure") or it.get("UnitOfMeasure")))
    for n, v in seen.items(): print("   ", n, v[:2])
