"""
Probe 1 for monthly installed capacity, group PL CZ SK HU RO BG AT CH.
Fetches candidate URLs from Actions, prints status / type / size / head, and for JSON/CSV/XLS tries to show structure.
Findings are recorded in the docstring of the follow-up probes and the build script.
Usage: python3 EU_CAPACITY_CEE_PROBE1.py
"""
import json, re, sys, time
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
URLS = [
 # Energy-Charts fallback
 ("EC pl monthly", "https://api.energy-charts.info/installed_power?country=pl&time_step=monthly&installation_decommission=false"),
 ("EC cz monthly", "https://api.energy-charts.info/installed_power?country=cz&time_step=monthly&installation_decommission=false"),
 ("EC sk monthly", "https://api.energy-charts.info/installed_power?country=sk&time_step=monthly&installation_decommission=false"),
 ("EC hu monthly", "https://api.energy-charts.info/installed_power?country=hu&time_step=monthly&installation_decommission=false"),
 ("EC ro monthly", "https://api.energy-charts.info/installed_power?country=ro&time_step=monthly&installation_decommission=false"),
 ("EC bg monthly", "https://api.energy-charts.info/installed_power?country=bg&time_step=monthly&installation_decommission=false"),
 ("EC at monthly", "https://api.energy-charts.info/installed_power?country=at&time_step=monthly&installation_decommission=false"),
 ("EC ch monthly", "https://api.energy-charts.info/installed_power?country=ch&time_step=monthly&installation_decommission=false"),
 # Poland
 ("PSE api root", "https://api.raporty.pse.pl/api/$metadata"),
 ("PSE pl installed", "https://www.pse.pl/dane-systemowe/funkcjonowanie-kse/raporty-roczne-z-funkcjonowania-kse-za-rok/moc-zainstalowana"),
 ("PSE pl bilans", "https://www.pse.pl/web/pse-eng/data/system-data"),
 ("ARE", "https://www.are.waw.pl/"),
 ("URE OZE rejestr", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/5753,Moc-zainstalowana-MW.html"),
 ("URE OZE potencjal", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze"),
 # Czechia
 ("ERU monthly", "https://www.eru.gov.cz/en/mesicni-zpravy-o-provozu-ees-cr"),
 ("ERU monthly cz", "https://www.eru.gov.cz/mesicni-zpravy-o-provozu-ees-cr"),
 ("ERU stat", "https://www.eru.gov.cz/en/statistiky-elektrina"),
 ("CEPS ", "https://www.ceps.cz/en/all-data"),
 ("CEPS inst", "https://www.ceps.cz/en/installed-capacity"),
 # Slovakia
 ("SEPS", "https://www.sepsas.sk/Dokumenty/Vykaz.asp?kod=24"),
 ("SEPS en", "https://www.sepsas.sk/en/"),
 ("URSO", "https://www.urso.gov.sk/"),
 # Hungary
 ("MAVIR data", "https://www.mavir.hu/web/mavir-en/installed-capacity"),
 ("MAVIR hu", "https://www.mavir.hu/web/mavir/beepitett-teljesitmeny"),
 ("MAVIR dataportal", "https://www.mavir.hu/web/mavir-en/data-portal"),
 ("MEKH", "https://www.mekh.hu/"),
 ("MAVIR VER", "https://www.mavir.hu/web/mavir/vet"),
 # Romania
 ("Transelectrica", "https://www.transelectrica.ro/en/web/tel/productie"),
 ("Transelectrica mix", "https://www.transelectrica.ro/widget/web/tel/sen-harta/-/harta_WAR_SENOperareHartaportlet"),
 ("Transelectrica pi", "https://www.transelectrica.ro/web/tel/putere-instalata"),
 ("ANRE", "https://www.anre.ro/"),
 ("ANRE rapoarte", "https://anre.ro/despre/rapoarte/rapoarte-de-monitorizare-a-pietei-de-energie-electrica/"),
 # Bulgaria
 ("ESO", "https://www.eso.bg/"),
 ("ESO en", "https://www.eso.bg/?lang=en"),
 ("ESO inst", "https://www.eso.bg/en/electricity-market/data"),
 ("EWRC", "https://www.dker.bg/en/"),
 ("EWRC reports", "https://www.dker.bg/en/reports"),
 # Austria
 ("Econtrol stat", "https://www.e-control.at/statistik/strom"),
 ("Econtrol bilanz", "https://www.e-control.at/statistik/e-statistik/archiv/betriebsstatistik"),
 ("Econtrol bestand", "https://www.e-control.at/statistik/e-statistik/archiv/bestandsstatistik"),
 ("APG", "https://markttransparenz.apg.at/"),
 ("APG installed", "https://markttransparenz.apg.at/en/markt/Markttransparenz/Erzeugung/Installierte-Leistung"),
 ("APG installed de", "https://markttransparenz.apg.at/de/markt/Markttransparenz/Erzeugung/Installierte-Leistung"),
 # Switzerland
 ("BFE ogd", "https://www.uvek-gis.admin.ch/BFE/ogd/"),
 ("BFE ogd104", "https://www.uvek-gis.admin.ch/BFE/ogd/104/ogd104_stromproduktion_swissgrid.csv"),
 ("BFE Pronovo OGD", "https://opendata.swiss/en/organization/bundesamt-fur-energie-bfe"),
 ("BFE ogd22", "https://www.uvek-gis.admin.ch/BFE/ogd/22/ogd22_energieerzeugungsanlagen.csv"),
 ("BFE ogd56", "https://www.uvek-gis.admin.ch/BFE/ogd/56/ogd56_wasserkraftanlagen.csv"),
 ("BFE ogd33", "https://www.uvek-gis.admin.ch/BFE/ogd/33/ogd33_pronovo.csv"),
 ("BFE ogd34", "https://www.uvek-gis.admin.ch/BFE/ogd/34/ogd34_pronovo.csv"),
 ("opendata.swiss ckan", "https://ckan.opendata.swiss/api/3/action/package_search?q=Elektrizit%C3%A4t%20installierte%20Leistung&rows=20"),
 ("opendata.swiss ckan2", "https://ckan.opendata.swiss/api/3/action/package_search?q=organization:bundesamt-fur-energie-bfe&rows=100"),
 ("Pronovo", "https://pronovo.ch/de/services/dashboard/"),
 ("Pronovo data", "https://pronovo.ch/de/services/open-data/"),
 ("Swissgrid", "https://www.swissgrid.ch/en/home/operation/grid-data/generation.html"),
 ("Swissolar", "https://www.uvek-gis.admin.ch/BFE/ogd/17/ogd17_fixpunkte_zur_heizkostenabrechnung.csv"),
]
def sq(t, n=260): return re.sub(r"\s+", " ", t)[:n]
for name, url in URLS:
    t0 = time.time()
    try:
        r = requests.get(url, headers=H, timeout=(10, 45))
        ct = r.headers.get("content-type", "")
        head = sq(r.text) if ("text" in ct or "json" in ct or "xml" in ct) else f"<{len(r.content)} bytes binary>"
        print(f"[{r.status_code}] {name} {round(time.time()-t0,1)}s {ct[:40]} len={len(r.content)}\n   {url}\n   {head}", flush=True)
        if "energy-charts" in url and r.status_code == 200:
            try:
                j = r.json(); print("   EC keys:", list(j.keys())[:6], "n_time", len(j.get("time", [])), [s["name"] for s in j.get("production_types", [])][:30])
                print("   first/last time:", j["time"][:2], j["time"][-2:])
            except Exception as e: print("   json err", e)
        if r.status_code == 200 and "html" in ct:
            links = sorted(set(re.findall(r'href=["\']([^"\']+\.(?:xlsx?|csv|pdf|zip|json))', r.text, re.I)))[:25]
            if links: print("   file links:", links)
    except Exception as e:
        print(f"[ERR] {name} {type(e).__name__}: {str(e)[:150]}\n   {url}", flush=True)
