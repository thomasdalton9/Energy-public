"""Round 5: ICC list-stat (operational data) and AEP get-database-data content (generation/capacity, West Africa)."""
import re, json, requests
from urllib.parse import urljoin
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) Chrome/124 Safari/537.36"
r = S.get("http://icc.ecowapp.org/list-stat", timeout=30); print("list-stat", r.status_code, len(r.content))
t = re.sub(r"\s+", " ", re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S))
print(re.sub(r"<[^>]+>", " ", t)[:1200])
for h, x in sorted(set((h, re.sub(r"<[^>]+>", "", x).strip()[:60]) for h, x in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.I | re.S))):
    if re.search(r"stat|xls|csv|pdf|indicator|data|export|json|chart", h + x, re.I): print("  ", urljoin(r.url, h), "|", x)
for m in sorted(set(re.findall(r'(?:src|href|data-url|action)="([^"]*(?:json|csv|xls|api|chart|highcharts)[^"]*)"', r.text, re.I)))[:20]: print("  attr", m)
print("=== AEP get-database-data")
d = S.get("https://africa-energy-portal.org/get-database-data", timeout=90)
print(d.status_code, len(d.content))
j = d.json(); print("indicator groups:", len(j))
rows = [(g["_id"], x) for g in j for x in g["data"]]
print("rows", len(rows))
gen = sorted({g["_id"] for g in j if re.search(r"generat|capacity|installed|hydro|import|export|electricity prod", g["_id"], re.I)})
for g in gen: print("  ind:", g)
WA = {"Nigeria","Ghana","Senegal","Cote d'Ivoire","Côte d'Ivoire","Mali","Guinea","Benin","Togo","Burkina Faso","Niger","Sierra Leone","Liberia","Gambia","The Gambia","Guinea-Bissau","Cabo Verde","Mauritania"}
for g in j:
    if g["_id"] in gen[:60]:
        ds = g["data"]; ctry = {x["name"] for x in ds if x["name"] in WA}; yrs = sorted({x["year"] for x in ds if x["name"] in WA})
        print(f"  {g['_id'][:70]} | src={ds[0].get('indicator_source')} | unit={ds[0].get('unit')} | WA countries={len(ctry)} | years {yrs[:1]}..{yrs[-1:]}")
print("sample row:", json.dumps(rows[0][1])[:300])
