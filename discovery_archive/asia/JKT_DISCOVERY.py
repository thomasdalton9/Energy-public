"""
Japan / Korea / Taiwan discovery round 1: which candidate endpoints answer from a GitHub runner, and what they return.
Prints status, content type, size and a short decoded head for each probe. Keys come from the environment (all
optional; probes that need one are skipped when it is not set):
  DATA_GO_KR_KEY  - Korea data.go.kr service key (KPX datasets)
  ESTAT_APP_ID    - Japan e-Stat appId (trade statistics)
  EMBER_API_KEY   - Ember
Run by .github/workflows/jkt_discovery.yml.
"""
import os
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
KEY = os.environ.get("DATA_GO_KR_KEY", "")
ESTAT = os.environ.get("ESTAT_APP_ID", "")


def head(raw, n=350):
    for enc in ("utf-8", "cp949", "cp932", "big5"):
        try:
            return raw[:n * 3].decode(enc)[:n].replace("\n", " | ").replace("\r", "")
        except UnicodeDecodeError:
            continue
    return repr(raw[:n])


def probe(label, url, params=None, verify=True):
    try:
        r = requests.get(url, params=params, headers=H, timeout=(15, 60), verify=verify)
        print(f"[{r.status_code}] {label}\n    {r.url[:160]}\n    {r.headers.get('content-type', '')} {len(r.content)} bytes\n"
              f"    {head(r.content)}\n", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"[ERR] {label}: {type(e).__name__}: {str(e)[:150]}\n", flush=True)


print("=== TAIWAN: Taipower open data ===")
for host in ("https://service.taipower.com.tw/data/opendata/apply/file", "https://data.taipower.com.tw/opendata/apply/file"):
    probe(f"live generation by unit ({host.split('/')[2]})", f"{host}/d006001/001.json")
probe("live genary", "https://www.taipower.com.tw/d006/loadGraph/loadGraph/data/genary.json")
probe("opendata catalogue", "https://data.taipower.com.tw/opendata/apply/case/")
for code in ("d006002", "d006003", "d006004", "d006005", "d006006", "d006007", "d006008", "d006009", "d006010", "d006011",
             "d006012", "d006013", "d006014", "d006015", "d006016", "d006017", "d006018"):
    for ext in ("csv", "json"):
        probe(f"{code}/001.{ext}", f"https://service.taipower.com.tw/data/opendata/apply/file/{code}/001.{ext}")

print("=== TAIWAN: MOEA / Bureau of Energy ===")
probe("E-STAT root", "https://ea01.moeaea.gov.tw/")
probe("energy statistics", "https://web3.moeaea.gov.tw/ECW/populace/web_book/WebReports.aspx?book=M_CH&menu_id=142")

print("=== KOREA: KPX ===")
probe("EPSIS home", "https://epsis.kpx.or.kr/epsisnew/")
probe("KPX open data", "https://openapi.kpx.or.kr/")
probe("KPX main", "https://www.kpx.or.kr/eng/")
if KEY:
    for path in ("SmpWithForecastDemand/getSmpWithForecastDemand", "PwrAmountByGen/getPwrAmountByGen",
                 "ElectricityDemand/getElectricityDemand", "SmpInfo/getSmpInfo"):
        probe(f"data.go.kr B552115/{path}", f"https://apis.data.go.kr/B552115/{path}",
              {"serviceKey": KEY, "pageNo": 1, "numOfRows": 5, "dataType": "json"})
    probe("data.go.kr KPX dataset search", "https://www.data.go.kr/tcs/dss/selectDataSetList.do",
          {"keyword": "한국전력거래소 발전량", "dType": "API"})
else:
    print("(DATA_GO_KR_KEY not set - KPX API probes skipped)\n")

print("=== KOREA: gas ===")
probe("KOGAS", "https://www.kogas.or.kr/site/koGas/ex/stats/")
probe("KESIS", "https://www.kesis.net/")
probe("KEEI energy stats", "https://www.keei.re.kr/")

print("=== JAPAN: TSO / JEPX / gas ===")
probe("JEPX spot FY2025", "https://www.jepx.jp/market/excel/spot_2025.csv")
probe("TEPCO PG area jukyu Apr-2025", "https://www.tepco.co.jp/forecast/html/images/eria_jukyu_202504_03.csv")
probe("OCCTO", "https://www.occto.or.jp/en/")
probe("METI LNG stocks", "https://www.meti.go.jp/statistics/tyo/denryoku_gas/index.html")
probe("METI gas statistics", "https://www.meti.go.jp/statistics/tyo/gas/index.html")
if ESTAT:
    probe("e-Stat statsList (trade stats)", "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsList",
          {"appId": ESTAT, "searchWord": "貿易統計 液化天然ガス", "limit": 10})
else:
    print("(ESTAT_APP_ID not set - e-Stat probes skipped)\n")
sys.exit(0)
