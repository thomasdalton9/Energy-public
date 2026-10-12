"""Probe 2: dig into the leads from probe 1 (FR ES PT IT monthly installed capacity).
Probe 1 findings (12 Oct 2026, Actions): Energy-Charts installed_power monthly = DE only; REData potencia-instalada with
time_trunc=month and /en/ gave HTTP 500; ODRE lists 15 'registre*' datasets incl. registre-national-installation-
production-stockage-electricite-agrege-311224; dgeg.gov.pt, terna.it, dati.terna.it reachable; gse.it 403 (Akamai)."""
import json, re, subprocess, sys
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=1500, show=True, ua=UA):
    o = subprocess.run(["curl", "-sL", "-m", "40", "-A", ua, "-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u],
                       capture_output=True, text=True, timeout=60)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u, "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

def links(html, pat, base=""):
    out = []
    for m in re.finditer(r'href="([^"]+)"', html):
        h = m.group(1)
        if re.search(pat, h, re.I):
            out.append(h)
    return sorted(set(out))

print("#### A. REData variants")
for q in ["generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month",
          "generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=year",
          "generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-03-31T23:59&time_trunc=month",
          "generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month&geo_trunc=electric_system&geo_limit=peninsular&geo_ids=8741",
          "generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=day"]:
    for lang in ("es", "en"):
        get(f"https://apidatos.ree.es/{lang}/datos/{q}", 600)

print("\n#### B. ODRE catalogue (parc / capacit / puissance / registre / installation)")
b = get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=100&where=search(dataset_id,%22parc%22)%20or%20search(dataset_id,%22capacite%22)%20or%20search(dataset_id,%22puissance%22)%20or%20search(dataset_id,%22installation%22)%20or%20search(dataset_id,%22registre%22)&select=dataset_id,records_count,data_processed", show=False)
try:
    for r in json.loads(b)["results"]:
        print(r)
except Exception as e:
    print("parse", e, b[:300])
for ds in ["registre-national-installation-production-stockage-electricite-agrege-311224", "parc-prod-par-filiere", "parc-national-annuel-prod-par-filiere", "parc-annuel-prod-par-filiere"]:
    m = get(f"https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{ds}", 200, show=True)
    try:
        j = json.loads(m)
        print("FIELDS", [(f["name"], f["type"]) for f in j.get("fields", [])])
        print("META", j["metas"]["default"].get("title"), j["metas"]["default"].get("data_processed"))
    except Exception:
        pass
    get(f"https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{ds}/records?limit=3", 1500)

print("\n#### C. DGEG links")
for u in ["https://www.dgeg.gov.pt/pt/estatistica/energia/eletricidade/", "https://dgeg.gov.pt/pt/estatistica/energia/estatisticas-rapidas/"]:
    h = get(u, show=False)
    print(u, len(h))
    for l in links(h, r"renov|potenc|xls|pdf|eletric|estatisticas-rapidas"):
        print("  ", l)

print("\n#### D. Terna")
for u in ["https://dati.terna.it/en/generation/installed-capacity", "https://www.terna.it/en/electric-system/statistics/statistical-publications",
          "https://www.terna.it/it/sistema-elettrico/statistiche/pubblicazioni-statistiche", "https://www.terna.it/it/sistema-elettrico/statistiche/rapporto-mensile"]:
    h = get(u, show=False)
    print(u, len(h))
    for l in links(h, r"\.xlsx|\.xls|\.pdf|download|capacity|potenza|api|rapporto|monthly|mensile|installed"):
        print("  ", l[:200])
    for m in re.finditer(r'(https?://[^"\' ]*(?:api|download|export)[^"\' ]*)', h):
        print("  URL", m.group(1)[:200])

print("\n#### E. REN datahub")
h = get("https://datahub.ren.pt/en/", show=False)
for l in links(h, r"capac|potenc|install|generation|production"):
    print("  ", l)
get("https://servicebus.ren.pt/datahubapi/electricity/InstalledCapacity", 400)
get("https://servicebus.ren.pt/datahubapi/electricity/ElectricityGenerationInstalledCapacity", 400)
get("https://servicebus.ren.pt/datahubapi/electricity/InstalledPower?culture=en", 400)

print("\n#### F. GSE / SDES (curl_cffi)")
try:
    from curl_cffi import requests as cr
    for u in ["https://www.gse.it/dati-e-scenari/statistiche", "https://www.gse.it/servizi-per-te/fotovoltaico/osservatorio",
              "https://www.statistiques.developpement-durable.gouv.fr/publications?rubrique=22&thematique=104",
              "https://www.statistiques.developpement-durable.gouv.fr/tableau-de-bord-solaire-photovoltaique-quatrieme-trimestre-2025",
              "https://www.statistiques.developpement-durable.gouv.fr/tableau-de-bord-eolien-troisieme-trimestre-2026",
              "https://www.statistiques.developpement-durable.gouv.fr/tableau-de-bord-solaire-photovoltaique-deuxieme-trimestre-2026"]:
        try:
            r = cr.get(u, impersonate="chrome", timeout=40)
            print("\n=====", r.status_code, len(r.content), u, flush=True)
            for l in links(r.text, r"\.xlsx|\.xls|\.csv|\.pdf|tableau-de-bord|parc|raccord|file"):
                print("  ", l[:200])
        except Exception as e:
            print("ERR", u, type(e).__name__, str(e)[:100])
except Exception as e:
    print("no curl_cffi", e)
