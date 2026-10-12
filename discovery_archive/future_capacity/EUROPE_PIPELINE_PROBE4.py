"""
Probe 4: official nuclear / closure / offshore documents - fetch each candidate page or PDF and print the text around key terms.
"""
import io, re, sys
import requests
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
try:
    import pdfplumber
except Exception: pdfplumber = None


def text(u):
    r = requests.get(u, headers=UA, timeout=(20, 120))
    ct = r.headers.get("content-type", "")
    print("\n#####", r.status_code, len(r.content), ct[:30], u[:160])
    if r.status_code != 200: return ""
    if "pdf" in ct or u.lower().split("?")[0].endswith(".pdf"):
        try:
            with pdfplumber.open(io.BytesIO(r.content)) as p:
                return "\n".join((pg.extract_text() or "") for pg in p.pages[:60])
        except Exception as e: print("pdf err", e); return ""
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))


def ctx(t, pats, w=160, n=4):
    for p in pats:
        ms = list(re.finditer(p, t, re.I))
        print(f"  [{p}] {len(ms)} hits")
        for m in ms[:n]:
            print("     ...", t[max(0, m.start() - w): m.end() + w].replace("\n", " "))


JOBS = [
 ("https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c", [r"3,?2\d\d ?MW", r"203\d", r"2029", r"unit 1"]),
 ("https://www.sizewellc.com/", [r"3,?2\d\d ?MW|3\.2 ?GW", r"203\d", r"final investment"]),
 ("https://www.paks2.hu/", [r"MW", r"20[23]\d"]),
 ("https://www.cez.cz/en/about-us/energy-mix/nuclear-power-plants/dukovany-new-build", [r"MW", r"203\d"]),
 ("https://www.ppej.pl/", [r"MW|GW", r"203\d"]),
 ("https://www.seas.sk/", [r"Mochovce", r"MW"]),
 ("https://www.nuclearelectrica.ro/", [r"Cernavoda|Cernavodă", r"MW"]),
 ("https://www.engie.com/en/", [r"Doel 4|Tihange 3", r"2035"]),
 ("https://www.ensi.ch/en/", [r"Beznau|Leibstadt"]),
 ("https://www.foronuclear.org/en/", [r"Almaraz"]),
 ("https://www.miteco.gob.es/es/energia/nuclear/calendario-cierre-centrales.html", [r"Almaraz", r"20[23]\d"]),
 ("https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Ausschreibungen/Wind_Offshore/start.html", [r"Zuschlag", r"MW|GW"]),
 ("https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Ausschreibungen/Kohleausstieg/start.html", [r"Steinkohle"]),
 ("https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Ausschreibungen/start.html", [r"Offshore|Kohle"]),
 ("https://www.bsh.de/DE/THEMEN/Offshore/Meeresfachplanung/meeresfachplanung_node.html", [r"Flächenentwicklungsplan|Flaechenentwicklungsplan"]),
 ("https://english.rvo.nl/topics/offshore-wind-energy", [r"GW|MW", r"203\d"]),
 ("https://www.rvo.nl/onderwerpen/windenergie-op-zee", [r"GW", r"203\d"]),
 ("https://ens.dk/en/our-responsibilities/offshore-wind-tenders", [r"GW|MW", r"20[23]\d"]),
 ("https://www.ecologie.gouv.fr/politiques-publiques/eolien-mer", [r"GW|MW", r"20[23]\d"]),
 ("https://www.gov.pl/web/climate/offshore-wind-energy", [r"GW|MW", r"20[23]\d"]),
 ("https://www.pse.pl/web/pse-eng/", [r"coal|capacity"]),
 ("https://www.terna.it/en/electric-system/grid-development/connection-requests", [r"GW|MW", r"20[23]\d"]),
]
for u, pats in JOBS:
    try:
        t = text(u)
        print("  chars", len(t)); ctx(t, pats)
    except Exception as e: print("ERR", u, repr(e)[:120])
    sys.stdout.flush()
