"""Probe 2 (12 Oct 2026): MaStR bulk download links, Elia dataset list, CBS capacity tables, Energy-Charts DE structure."""
import re, json, requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"
def j(u, **kw):
    r = S.get(u, timeout=90, **kw); return r
print("### MaStR Datendownload links")
r = j("https://www.marktstammdatenregister.de/MaStR/Datendownload")
t = r.text
for m in re.finditer(r'(https?://[^"\' <>]+\.(?:zip|xml|csv)[^"\' <>]*|href="[^"]*(?:ownload|\.zip)[^"]*")', t):
    print("  ", m.group(0)[:250])
i = t.find("Gesamtdaten")
print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t[max(0, i-500):i+3000]))[:2500])
print("### Elia dataset list")
off = 0
while True:
    r = j(f"https://opendata.elia.be/api/explore/v2.1/catalog/datasets?limit=100&offset={off}&select=dataset_id,title")
    d = r.json()
    for x in d["results"]:
        print("  ", x["dataset_id"], "|", x["title"][:110])
    off += 100
    if off >= d["total_count"]: break
print("### CBS tables with capacity words")
for w in ["vermogen", "capaciteit", "Zonne", "Wind", "Hernieuwbare", "Elektriciteit"]:
    r = j("https://opendata.cbs.nl/ODataCatalog/Tables?$format=json&$select=Identifier,Title,Frequency,Period,Updated&$filter=substringof('%s',Title)" % w)
    for x in r.json()["value"]:
        print("  ", w, "|", x["Identifier"], "|", x["Title"][:100], "|", x["Frequency"], "|", x["Period"], "|", x["Updated"][:10])
print("### Energy-Charts DE")
r = j("https://api.energy-charts.info/installed_power?country=de&time_step=monthly&installation_decommission=false")
d = r.json()
print(list(d.keys()), len(d["time"]), d["time"][-3:])
for s in d["production_types"]:
    print("  ", s["name"], s["data"][-3:] if isinstance(s["data"], list) else s["data"])
for c in ["nl", "be", "lu"]:
    r = j(f"https://api.energy-charts.info/installed_power?country={c}&time_step=yearly&installation_decommission=false")
    print(c, r.status_code, r.text[:200])
    try:
        d = r.json(); print(d["time"][-3:], [s["name"] for s in d["production_types"]])
    except Exception as e: print(e)
