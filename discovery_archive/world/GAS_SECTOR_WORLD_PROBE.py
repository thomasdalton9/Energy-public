"""
Probe candidate raw sources for gas demand by sector outside the Americas/ANZ/Singapore/Ireland.
Prints status / content type / first bytes per URL so the survey report
(discovery_archive/world/GAS_DEMAND_BY_SECTOR_WORLD_SURVEY.md) can say what is reachable from the
Claude sandbox (via the agent proxy) vs. what needs GitHub Actions.

Usage: python3 GAS_SECTOR_WORLD_PROBE.py [label_filter]
"""
import sys

import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
ES = "https://ec.europa.eu/eurostat/api/dissemination"
URLS = {
    "eurostat nrg_cb_gasm dataflow": f"{ES}/sdmx/2.1/dataflow/ESTAT/nrg_cb_gasm/1.0",
    "eurostat nrg_cb_gasm json DE": f"{ES}/statistics/1.0/data/nrg_cb_gasm?format=JSON&lang=en&geo=DE&unit=TJ_GCV&sinceTimePeriod=2025-01",
    "eurostat nrg_cb_gasm csv DE": f"{ES}/sdmx/2.1/data/nrg_cb_gasm/M..?geo=DE&startPeriod=2025-01&format=SDMX-CSV",
    "eurostat nrg_cb_gasm dsd": f"{ES}/sdmx/2.1/datastructure/ESTAT/NRG_CB_GASM/latest?references=descendants&detail=referencepartial",
    "eurostat nrg_bal_c annual": f"{ES}/statistics/1.0/data/nrg_bal_c?format=JSON&geo=DE&siec=G3000&unit=TJ&time=2023&nrg_bal=FC_IND_E",
    "eurostat nrg_cb_gas annual": f"{ES}/statistics/1.0/data/nrg_cb_gas?format=JSON&geo=DE&unit=TJ_GCV&time=2023",
    "entsog operationalData": "https://transparency.entsog.eu/api/v1/operationalData?limit=1&indicator=Physical%20Flow&periodType=day",
    "entsog aggregatedData": "https://transparency.entsog.eu/api/v1/AggregatedData?limit=2",
    "jodi gas csv": "https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv_beta.zip",
    "jodi gas page": "https://www.jodidata.org/gas/database/data-downloads.aspx",
    "ppac india gas consumption": "https://ppac.gov.in/natural-gas/consumption",
    "ppac india root": "https://ppac.gov.in/",
    "india mopng": "https://mopng.gov.in/en/petroleum-statistics/",
    "japan meti gas stats": "https://www.meti.go.jp/english/statistics/tyo/gas/index.html",
    "japan meti gas jp monthly": "https://www.meti.go.jp/statistics/tyo/gasdoukou/index.html",
    "japan meti LNG power": "https://www.meti.go.jp/statistics/tyo/denryoku_gas/index.html",
    "japan estat": "https://www.e-stat.go.jp/en",
    "japan JOGMEC": "https://www.jogmec.go.jp/english/",
    "korea kogas stats": "https://www.kogas.or.kr/portal/contents.do?key=1862",
    "korea data.go.kr": "https://www.data.go.kr/",
    "korea kesis": "https://www.kesis.net/",
    "china nbs data": "https://data.stats.gov.cn/english/",
    "china nbs release": "https://www.stats.gov.cn/english/PressRelease/",
    "taiwan moeaea": "https://www.moeaea.gov.tw/ECW/english/content/ContentLink.aspx?menu_id=1540",
    "taiwan esist": "https://www.esist.org.tw/",
    "thailand eppo gas": "https://www.eppo.go.th/index.php/en/en-energystatistics/ng-statistic",
    "thailand eppo root": "https://www.eppo.go.th/",
    "pakistan ogra": "https://www.ogra.org.pk/",
    "pakistan hdip": "https://hdip.com.pk/",
    "bangladesh petrobangla": "https://petrobangla.org.bd/",
    "uk desnz energy trends gas": "https://www.gov.uk/government/statistics/gas-section-4-energy-trends",
    "uk gov.uk content api gas": "https://www.gov.uk/api/content/government/statistics/gas-section-4-energy-trends",
    "uk NESO data portal": "https://api.neso.energy/api/3/action/package_search?q=gas",
    "uk national gas data": "https://data.nationalgas.com/",
    "germany BNetzA gas": "https://www.bundesnetzagentur.de/DE/Gasversorgung/aktuelle_gasversorgung/_svg/Gasverbrauch/Gasverbrauch.html",
    "germany THE": "https://www.tradinghub.eu/en-gb/Publications/Transparency",
    "germany destatis genesis": "https://www-genesis.destatis.de/",
    "france odre catalog": "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets?limit=3&where=search(%22consommation%22)",
    "france odre conso brute": "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/consommation-quotidienne-brute/records?limit=2",
    "italy snam": "https://www.snam.it/en/transport/data-and-information/",
    "netherlands CBS odata": "https://opendata.cbs.nl/ODataApi/odata/00372/TableInfos",
    "belgium Fluxys": "https://www.fluxys.com/en/",
    "norway ssb": "https://data.ssb.no/api/v0/en/table/",
    "norway gassco umm": "https://umm.gassco.no/",
    "turkey botas": "https://www.botas.gov.tr/",
    "turkey epias": "https://seffaflik.epias.com.tr/",
    "turkey epdk": "https://www.epdk.gov.tr/",
    "russia rosstat": "https://rosstat.gov.ru/",
    "kazakhstan stat": "https://stat.gov.kz/",
    "iran nigc": "https://www.nigc.ir/",
    "saudi gastat": "https://www.stats.gov.sa/",
    "uae fcsc": "https://fcsc.gov.ae/",
    "egypt capmas": "https://www.capmas.gov.eg/",
    "egypt petroleum": "https://www.petroleum.gov.eg/",
    "nigeria nuprc": "https://www.nuprc.gov.ng/",
    "algeria sonelgaz": "https://www.sonelgaz.dz/",
    "south africa energy": "https://www.energy.gov.za/",
    "ember data": "https://ember-energy.org/data/",
    "energy institute": "https://www.energyinst.org/statistical-review",
    "iea data": "https://www.iea.org/data-and-statistics",
    "owid energy csv": "https://raw.githubusercontent.com/owid/energy-data/master/owid-energy-data.csv",
}


def main():
    flt = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    for label, url in URLS.items():
        if flt and flt not in label.lower():
            continue
        try:
            r = requests.get(url, headers=UA, timeout=(8, 25), stream=True)
            body = next(r.iter_content(400), b"")
            r.close()
            print(f"{r.status_code:>4} {label:32} {r.headers.get('content-type','')[:30]:30} {body[:100]!r}", flush=True)
        except Exception as e:
            print(f" ERR {label:32} {type(e).__name__}: {str(e)[:100]}", flush=True)


if __name__ == "__main__":
    main()
