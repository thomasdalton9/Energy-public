"""Biomethane probe 3: ODRE France dataset heads + annual sums, Eurostat annual biogases, Germany sources. Prints only."""
import json, re, requests
import pandas as pd
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
def g(url, **kw):
    kw.setdefault("headers", UA)
    try:
        return requests.get(url, timeout=60, **kw)
    except Exception as e:
        print("ERR", url, type(e).__name__, e)
B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
print("=== ODRE")
for ds in ["production-definitive-et-journaliere-de-biomethane-tout-reseau", "odre-prod-grdgrt-operateur-def", "prod-def-reg-jour-biom-reseau-grtgrd"]:
    r = g(B + ds)
    j = r.json(); print("##", ds, "fields:", [(f["name"], f["type"]) for f in j["fields"]])
    dcol = [f["name"] for f in j["fields"] if f["type"] == "date"][0]
    r = g(B + ds + "/records", params={"limit": 2, "order_by": dcol + " desc"}); print("  last:", json.dumps(r.json().get("results"))[:700])
    r = g(B + ds + "/records", params={"limit": 2, "order_by": dcol + " asc"}); print("  first:", json.dumps(r.json().get("results"))[:700])
    val = [f["name"] for f in j["fields"] if f["type"] == "double" and re.search("prod|biom|mwh|gwh", f["name"], re.I)]
    for v in val[:3]:
        r = g(B + ds + "/records", params={"select": f"year({dcol}) as y, sum({v}) as s, count(*) as n", "group_by": f"year({dcol})", "order_by": "y", "limit": 30})
        print("  annual sum of", v, [(x["y"], round(x["s"] or 0), x["n"]) for x in r.json().get("results", [])] if r.ok else r.text[:200])
    for cat in [f["name"] for f in j["fields"] if f["type"] == "text" and f["name"] in ("statut", "operateur", "type_de_reseau", "operateur_de_transport", "reseau")]:
        r = g(B + ds + "/records", params={"select": f"{cat}, count(*) as n", "group_by": cat, "limit": 20}); print("  ", cat, [(x[cat], x["n"]) for x in r.json().get("results", [])])
print("=== Eurostat annual biogases")
for ds, par in [("nrg_bal_c", {"siec": ["R5300", "G3000"], "nrg_bal": ["IPRD", "GIC"], "unit": "TJ"}), ("nrg_cb_gas", {})]:
    p = {"geo": ["DE", "FR", "NL", "DK", "IE"], "sinceTimePeriod": "2020", "format": "JSON", "lang": "EN", **par}
    r = g(f"https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/{ds}", params=p)
    print(ds, r and r.status_code)
    if r is not None and r.ok:
        j = r.json(); dims = j["id"]; print(dims, {k: list(j["dimension"][k]["category"]["label"].items())[:25] for k in dims if k in ("siec", "nrg_bal") and len(j["dimension"][k]["category"]["label"]) < 80})
        idx = {k: list(j["dimension"][k]["category"]["index"]) for k in dims}
        import itertools
        sizes = j["size"]
        for pos, v in j["value"].items():
            pos = int(pos); coords = []
            for s, k in reversed(list(zip(sizes, dims))):
                coords.append(idx[k][pos % s]); pos //= s
            print("  ", dict(zip(reversed(dims), coords)) if False else list(reversed(coords)), v)
    else:
        print(r and r.text[:300])
print("=== Germany")
for u in ["https://www.biogaspartner.de/biogaspartner/biomethan/einspeiseatlas/", "https://www.biogaspartner.de/biogaspartner/infocenter/", "https://www.fachverband-biogas.de/presse/"]:
    r = g(u); print(u, r and r.status_code, r and len(r.text))
    if r is not None and r.ok:
        t = re.sub(r"<[^>]+>", " ", r.text)
        for m in re.finditer(r"[^.]{0,120}(TWh|Mrd\. ?kWh|Milliarden)[^.]{0,120}", t): print("   ", re.sub(r"\s+", " ", m.group(0))[:250])
        print("  links:", sorted(set(re.findall(r'href="([^"]*(?:xls|csv|pdf|statist)[^"]*)"', r.text, re.I)))[:15])
r = g("https://www.destatis.de/DE/Themen/Branchen-Unternehmen/Energie/Erzeugung/_inhalt.html"); print("destatis", r and r.status_code)
if r is not None and r.ok: print(sorted(set(re.findall(r'href="([^"]*(?:gas|Gas)[^"]*)"', r.text)))[:20])
r = g("https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Monitoringberichte/start.html")
if r is not None and r.ok: print(sorted(set(re.findall(r'href="([^"]*Monitoring[^"]*)"', r.text)))[:15])
r = g("https://www.marktstammdatenregister.de/MaStR/Datendownload"); print("MaStR", r and r.status_code)
