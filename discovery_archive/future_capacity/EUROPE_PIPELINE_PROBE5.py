"""
Probe 5 (compact): BNetzA Zu- und Rueckbau workbook layout, REPD columns, ODRE register/queue datasets, gov.uk CfD rounds and DUKES
attachments, ENTSO-E ERAA download pages, Commission NECP file names.
"""
import io, json, re
import pandas as pd
import requests
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}


def get(u, **kw):
    r = requests.get(u, headers=UA, timeout=(20, 120), **kw)
    print("GET", r.status_code, len(r.content), u[:140]); return r


def sec(t): print("\n=========", t, flush=True)


def show(d, rows=None, w=26):
    for i, row in d.iterrows():
        if rows and i not in rows: continue
        cells = ["" if pd.isna(v) else str(v).replace("\n", " ")[:w] for v in row.tolist()]
        if any(cells): print(i, "|".join(cells))


sec("BNetzA")
b = "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/_DL/"
for f in ["ZuUndRueckbau.xlsx?__blob=publicationFile&v=12", "Kraftwerksliste.xlsx?__blob=publicationFile&v=17"]:
    x = pd.ExcelFile(io.BytesIO(get(b + f).content))
    for s in x.sheet_names:
        d = x.parse(s, header=None)
        print("--- sheet", s, d.shape)
        if "einzeln" in s:
            show(d, w=22)
        else:
            show(d.head(40), w=30)

sec("REPD columns")
r = get("https://assets.publishing.service.gov.uk/media/6a6cbdc00c36759b5ccaa305/REPD_Publication_Q2_2026.csv")
d = pd.read_csv(io.BytesIO(r.content), encoding="cp1252", low_memory=False)
print(list(d.columns))
ex = d[d["Development Status (short)"] == "Under Construction"]
print(ex.groupby("Technology Type")["Installed Capacity (MWelec)"].agg(["count", lambda s: pd.to_numeric(s, errors="coerce").sum()]))
aw = d[d["Development Status (short)"] == "Awaiting Construction"]
print(aw.groupby("Technology Type")["Installed Capacity (MWelec)"].agg(["count", lambda s: pd.to_numeric(s, errors="coerce").sum()]))
print(d[(d["Technology Type"] == "Wind Offshore") & d["Development Status (short)"].isin(["Under Construction", "Awaiting Construction"])][["Site Name", "Installed Capacity (MWelec)", "Development Status (short)", "CfD Allocation Round", "Offshore Wind Round", "Under Construction"]].to_string())

sec("ODRE")
j = get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=100&where=search(%22registre%22)").json()
for res in j["results"]:
    print(res["dataset_id"], "|", (res["metas"]["default"].get("title") or "")[:90], "|", res["metas"]["default"].get("records_count"))
for q in ["projet", "raccordement", "capacit", "planifi", "mise%20en%20service"]:
    j = get(f"https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=40&where=search(%22{q}%22)").json()
    for res in j.get("results", []):
        print(q, res["dataset_id"], "|", (res["metas"]["default"].get("title") or "")[:90])

sec("gov.uk CfD")
for ar in ["4", "5", "6", "7"]:
    try:
        c = get(f"https://www.gov.uk/api/content/government/publications/contracts-for-difference-cfd-allocation-round-{ar}-results").json()
        for a in c.get("details", {}).get("attachments", []): print("  AR", ar, a.get("title"), "|", a.get("url"))
        print("  docs", c.get("links", {}).get("documents", [{}])[0] if False else "")
    except Exception as e: print("AR", ar, repr(e)[:100])
j = get("https://www.gov.uk/api/search.json?q=allocation+round+results+contracts+for+difference&count=40&fields=title&fields=link").json()
for x in j["results"]: print(x["link"], "|", x["title"][:80])

sec("DUKES ch5")
c = get("https://www.gov.uk/api/content/government/statistics/electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes").json()
for a in c.get("details", {}).get("attachments", []): print(a.get("title"), "|", a.get("url"))

sec("ERAA")
for u in ["https://www.entsoe.eu/eraa/2025/", "https://www.entsoe.eu/eraa/2025/downloads", "https://www.entsoe.eu/eraa/2025/modelling-data", "https://www.entsoe.eu/eraa/2023/modelling-data", "https://www.entsoe.eu/eraa/2021/downloads", "https://www.entsoe.eu/outlooks/eraa/2025/"]:
    try:
        h = get(u).text
        for l in sorted(set(re.findall(r'href="([^"]+)"', h))):
            if re.search(r"(?i)\.(xlsx|csv|zip|xls)|download|national|capacit|dashboard", l): print("   ", l[:200])
    except Exception as e: print("ERR", u, repr(e)[:100])

sec("NECP")
h = get("https://commission.europa.eu/energy-climate-change-environment/implementation-eu-countries/energy-and-climate-governance-and-reporting/national-energy-and-climate-plans_en").text
for m in re.finditer(r'href="([^"]*filename=([^"&]+)[^"]*)"', h): print(m.group(2)[:110])
