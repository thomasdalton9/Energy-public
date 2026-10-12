"""Probe 3: follow-up to probe 2 (FR ES PT IT monthly installed capacity).
Probe 2 findings: REData potencia-instalada returned HTTP 500 for month/year/day, es/en (no Accept header sent);
ODRE has the national production register 'registre-national-installation-production-stockage-electricite-agrege' (current, 140,539 rows,
processed 2026-09-10) plus year-end snapshots 311217-311225, and 'parc-prod-par-filiere' (annual 19 rows); SDES publishes quarterly dashboards
'tableau-de-bord-solaire-photovoltaique-<trimestre>-<year>' and eolien (HTML pages, data links not yet found); DGEG has
'potencia-instalada-fer' and 'estatisticas-rapidas-das-renovaveis' pages; Terna pages list download.terna.it files; GSE is 403; REN servicebus
guesses 404."""
import json, re, subprocess
from html import unescape
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=()):
    cmd = ["curl", "-sL", "-m", "40", "-A", UA] + [x for h in hdr for x in ("-H", h)] + ["-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u, "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

def anchors(html, pat):
    for m in re.finditer(r'<a [^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        h, tx = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        tx = " ".join(unescape(tx).split())
        if re.search(pat, h + " " + tx, re.I):
            print("   ", h[:170], "|", tx[:100])

print("#### A. REData with Accept header; sanity endpoint")
H = ("Accept: application/json", "Content-Type: application/json")
get("https://apidatos.ree.es/en/datos/generacion/estructura-generacion?start_date=2025-01-01T00:00&end_date=2025-01-31T23:59&time_trunc=month", 500, hdr=H)
for q in ["time_trunc=month", "time_trunc=year"]:
    get(f"https://apidatos.ree.es/en/datos/generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&{q}", 1500, hdr=H)
get("https://apidatos.ree.es/es/datos/generacion/potencia-instalada?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month&geo_trunc=electric_system&geo_limit=peninsular&geo_ids=8741", 800, hdr=H)
get("https://apidatos.ree.es/es/datos/generacion/potencia-instalada?start_date=2021-01-01T00:00&end_date=2021-12-31T23:59&time_trunc=month", 800, hdr=H)

print("\n#### B. ODRE register fields and sample")
j = json.loads(get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/registre-national-installation-production-stockage-electricite-agrege", show=False))
print([(f["name"], f["type"]) for f in j["fields"]])
print(json.dumps(j["metas"]["default"], ensure_ascii=False)[:600])
get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/registre-national-installation-production-stockage-electricite-agrege/records?limit=2", 3500)
get("https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/registre-national-installation-production-stockage-electricite-agrege/exports/csv?limit=3&delimiter=%3B", 1500)

print("\n#### C. SDES quarterly pages")
for u in ["https://www.statistiques.developpement-durable.gouv.fr/tableau-de-bord-solaire-photovoltaique-deuxieme-trimestre-2026",
          "https://www.statistiques.developpement-durable.gouv.fr/tableau-de-bord-eolien-deuxieme-trimestre-2026"]:
    h = get(u, show=False)
    print(u, len(h))
    anchors(h, r"download|\.xls|\.csv|\.pdf|donn|data|fichier|telecharg|media|sites/default/files")
    for m in re.finditer(r"(?:puissance[^.<]{0,200}MW[^.<]{0,100})", unescape(re.sub(r"<[^>]+>", " ", h))):
        print("   TXT", " ".join(m.group(0).split())[:260])
        break

print("\n#### D. DGEG")
for u in ["https://www.dgeg.gov.pt/pt/estatistica/energia/renovaveis/potencia-instalada-fer/",
          "https://www.dgeg.gov.pt/pt/estatistica/energia/renovaveis/estatisticas-rapidas-das-renovaveis/",
          "https://www.dgeg.gov.pt/pt/estatistica/energia/eletricidade/producao-anual-e-potencia-instalada/",
          "https://www.dgeg.gov.pt/pt/estatistica/energia/eletricidade/producao-mensal-de-eletricidade/",
          "https://www.dgeg.gov.pt/pt/destaques/estatisticas-rapidas-das-renovaveis-agosto-2026/"]:
    h = get(u, show=False)
    print(u, len(h))
    anchors(h, r"media/|\.xls|\.csv|\.pdf")

print("\n#### E. Terna download center and pages")
for u in ["https://dati.terna.it/en/download-center", "https://www.terna.it/it/sistema-elettrico/statistiche/rapporto-mensile",
          "https://www.terna.it/en/electric-system/statistics/statistical-publications"]:
    h = get(u, show=False)
    print(u, len(h))
    anchors(h, r"download\.terna|\.xls|\.csv|\.pdf|installed|potenza|capacity|capacit")
    names = sorted(set(re.findall(r"download\.terna\.it/terna/[^\"'\\\s<>]+", h)))
    print("   FILES", names[:60])

print("\n#### F. REN centro de informacao")
for u in ["https://www.centrodeinformacao.ren.pt/", "https://www.centrodeinformacao.ren.pt/PT/InformacaoExploracao/Paginas/EstatisticaMensal.aspx",
          "https://datahub.ren.pt/en/homepage/electricity/production/detailed/"]:
    h = get(u, show=False)
    print(u, len(h))
    anchors(h, r"\.xls|\.pdf|potencia|installed|capacidade|capacity|mensal|monthly")

print("\n#### G. Energy-Charts yearly")
for c in ("fr", "es", "pt", "it"):
    get(f"https://api.energy-charts.info/installed_power?country={c}&time_step=yearly&installation_decommission=false", 300)
