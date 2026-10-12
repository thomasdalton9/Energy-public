"""
Probe 1 for europe/future/EUROPE_POWER_PIPELINE.py: which candidate pipeline sources answer from GitHub Actions.
Prints status / type / size / a text head for each URL, and for HTML pages the document links found.
Findings are recorded in the main script's 'Tried' sheet. Run: python3 EUROPE_PIPELINE_PROBE1.py
"""
import re
import sys
import requests

UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
URLS = [
    # UK
    "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract",
    "https://api.neso.energy/api/3/action/package_show?id=transmission-entry-capacity-tec-register",
    "https://www.neso.energy/data-portal/transmission-entry-capacity-tec-register",
    "https://www.gov.uk/api/search.json?q=contracts+for+difference+allocation+round+results&count=15&fields=title&fields=link",
    "https://www.gov.uk/api/content/government/statistics/digest-of-uk-energy-statistics-dukes-chapter-5-electricity",
    "https://www.gov.uk/api/content/government/statistics/electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes",
    # Germany
    "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/start.html",
    "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/DL_Kraftwerksliste.html",
    "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Ausschreibungen/Kohleausstieg/start.html",
    "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Ausschreibungen/Wind_Offshore/BesondereAusschreibungen/start.html",
    "https://www.bsh.de/DE/THEMEN/Offshore/Meeresfachplanung/Flaechenentwicklungsplan/flaechenentwicklungsplan_node.html",
    "https://www.marktstammdatenregister.de/MaStR/Datendownload",
    # France
    "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=20&where=search(%22file%20d%27attente%22)",
    "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=20&where=search(%22registre%22)",
    "https://data.rte-france.com/",
    "https://www.legifrance.gouv.fr/",
    # Spain
    "https://www.ree.es/es/clientes/generador/acceso-conexion/conoce-el-estado-de-tu-solicitud",
    "https://www.ree.es/es/clientes/generador/acceso-conexion",
    "https://www.miteco.gob.es/es/energia/",
    "https://www.csn.es/",
    # Italy / NL / BE / DK / PL / others
    "https://www.terna.it/en/electric-system/grid-development/connection-requests",
    "https://www.rvo.nl/onderwerpen/offshore-wind",
    "https://www.tennet.eu/",
    "https://opendata.elia.be/api/explore/v2.1/catalog/datasets?limit=30&where=search(%22capacity%22)",
    "https://www.pse.pl/",
    "https://www.ure.gov.pl/",
    "https://www.energinet.dk/",
    "https://api.energidataservice.dk/meta/dataset?limit=5",
    "https://www.entsoe.eu/outlooks/eraa/",
    "https://eepublicdownloads.entsoe.eu/",
    "https://www.iaea.org/",
    "https://pris.iaea.org/PRIS/CountryStatistics/CountryDetails.aspx?current=GB",
    "https://www.world-nuclear.org/information-library/country-profiles/countries-g-n/united-kingdom",
    "https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c",
    "https://www.gov.uk/government/publications/sizewell-c-final-investment-decision",
    "https://www.paks2.hu/",
    "https://commission.europa.eu/energy-climate-change-environment/implementation-eu-countries/energy-and-climate-governance-and-reporting/national-energy-and-climate-plans_en",
    "https://www.energy-charts.info/",
    "https://api.energy-charts.info/installed_power?country=de&time_step=yearly",
]


def main():
    for u in URLS:
        try:
            r = requests.get(u, headers=UA, timeout=(15, 40))
            ct = r.headers.get("content-type", "")
            txt = r.text[:200].replace("\n", " ") if "text" in ct or "json" in ct else ""
            print(f"{r.status_code} {len(r.content):>9} {ct[:40]:40} {u}\n      {txt}")
            if r.status_code == 200 and "html" in ct:
                links = sorted(set(re.findall(r'href="([^"]+\.(?:xlsx|xls|csv|pdf|zip|ods)[^"]*)"', r.text, re.I)))
                for l in links[:12]:
                    print("      LINK", l)
        except Exception as e:  # noqa: BLE001
            print("ERR", repr(e)[:100], u)
        sys.stdout.flush()


main()
