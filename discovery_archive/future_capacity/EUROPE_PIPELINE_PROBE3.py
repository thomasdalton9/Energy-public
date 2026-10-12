"""
Probe 3: structure of the machine-readable pipeline sources found by probe 1: BNetzA Kraftwerksliste / Zu- und Rueckbau, REPD,
NESO TEC register, ODRE queue dataset, gov.uk CfD allocation rounds, DUKES chapter 5, ENTSO-E ERAA links, Commission NECP links.
"""
import io, json, re, sys
import pandas as pd
import requests
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 60); pd.set_option("display.max_colwidth", 40)


def get(u, **kw):
    r = requests.get(u, headers=UA, timeout=(20, 120), **kw)
    print("GET", r.status_code, len(r.content), u[:150]); return r


def sec(t): print("\n=========", t, flush=True)


sec("BNetzA")
base = "https://www.bundesnetzagentur.de"
for p in ["/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/_DL/Kraftwerksliste.xlsx?__blob=publicationFile&v=17",
          "/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/_DL/ZuUndRueckbau.xlsx?__blob=publicationFile&v=12"]:
    try:
        r = get(base + p)
        x = pd.ExcelFile(io.BytesIO(r.content))
        for s in x.sheet_names:
            d = x.parse(s, header=None)
            print("--- sheet", s, d.shape); print(d.head(14).to_string())
    except Exception as e: print("ERR", repr(e)[:200])
try:
    html = get(base + "/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/start.html").text
    print(re.sub(r"<[^>]+>", " ", html)[3000:6000])
except Exception as e: print("ERR", e)

sec("REPD")
r = get("https://assets.publishing.service.gov.uk/media/6a6cbdc00c36759b5ccaa305/REPD_Publication_Q2_2026.csv")
d = pd.read_csv(io.BytesIO(r.content), encoding="cp1252", low_memory=False)
print(list(d.columns))
print(d["Development Status (short)"].value_counts())
print(d["Development Status"].value_counts())
print(d.groupby(["Development Status (short)", "Technology Type"])["Installed Capacity (MWelec)"].agg(["count"]).head(80))
uc = d[d["Development Status (short)"].astype(str).str.contains("Construction")]
print(uc.head(8).T.to_string())

sec("TEC")
j = get("https://api.neso.energy/api/3/action/package_show?id=transmission-entry-capacity-tec-register").json()
for res in j["result"]["resources"]: print(res["name"], res["format"], res["url"])
url = [x["url"] for x in j["result"]["resources"] if x["format"].upper() == "CSV"][0]
t = pd.read_csv(io.BytesIO(get(url).content), low_memory=False)
print(t.shape, list(t.columns))
for c in t.columns:
    if re.search(r"(?i)status|type|fuel|tech", c): print(c, t[c].value_counts().head(25).to_dict())
print(t.head(5).T.to_string())

sec("ODRE queue")
j = get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=20&where=search(%22file%20d%27attente%22)").json()
for res in j["results"]:
    print(res["dataset_id"], res["metas"]["default"].get("title"), res["metas"]["default"].get("records_count"))
    print([f["name"] for f in res["fields"]])
    rr = get(f"https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{res['dataset_id']}/records?limit=3").json()
    print(json.dumps(rr["results"], ensure_ascii=False)[:1500])
j = get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=100&where=search(%22raccordement%22%20OR%20%22installations%22%20OR%20%22projet%22)&select=dataset_id%2Cmetas").json()
for res in j["results"]: print(res["dataset_id"], "|", res["metas"]["default"].get("title"))

sec("gov.uk CfD")
j = get("https://www.gov.uk/api/search.json?q=contracts+for+difference+allocation+round+results&count=30&fields=title&fields=link").json()
for x in j["results"]: print(x["link"], "|", x["title"])
for ar in ["4", "5", "6", "7"]:
    try:
        c = get(f"https://www.gov.uk/api/content/government/publications/contracts-for-difference-cfd-allocation-round-{ar}-results").json()
        for a in c.get("details", {}).get("attachments", []): print("  AR", ar, a.get("title"), a.get("url"))
    except Exception as e: print("AR", ar, repr(e)[:100])

sec("DUKES ch5")
c = get("https://www.gov.uk/api/content/government/statistics/electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes").json()
for a in c.get("details", {}).get("attachments", []): print(a.get("title"), a.get("url"))

sec("ERAA")
h = get("https://www.entsoe.eu/outlooks/eraa/").text
for l in sorted(set(re.findall(r'href="([^"]+)"', h))):
    if re.search(r"(?i)eraa|resource-adequacy|national", l): print(l)
for u in ["https://www.entsoe.eu/outlooks/eraa/2025/", "https://eraa.entsoe.eu/", "https://www.entsoe.eu/outlooks/eraa/2024/", "https://tyndp.entsoe.eu/"]:
    try:
        h = get(u).text
        for l in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx|csv|zip|pdf)[^"]*)"', h)))[:40]: print("  ", l)
    except Exception as e: print("ERR", u, repr(e)[:100])

sec("NECP")
h = get("https://commission.europa.eu/energy-climate-change-environment/implementation-eu-countries/energy-and-climate-governance-and-reporting/national-energy-and-climate-plans_en").text
for m in re.finditer(r'href="([^"]*filename=[^"]+)"', h): print(m.group(1)[:200])
