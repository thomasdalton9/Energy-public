"""
Probe 13 (manual workflow south_korea_probe.yml): KESIS (national energy statistics, KEEI/MOTIE) statistics table ajax -
save the menu page and try the stat/list endpoints; data.go.kr dataset titles for the LNG / KOGAS search hits.
Everything saved to discovery_archive/results/south_korea/.
"""
import os
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9", "X-Requested-With": "XMLHttpRequest"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers.update(H)
K = "https://www.kesis.net"
r = S.get(K + "/menu.es?mid=a10101000000", timeout=(10, 60))
open(os.path.join(OUT, "kesis_menu_a10101000000.html"), "w", encoding="utf-8").write(r.text)
log = []
for name in ("selectStatsSrvcTblList", "getStatsInfo", "getStatsListInfo", "searchStats"):
    for data in ({}, {"mid": "a10101000000"}, {"mid": "a10101000000", "statsId": "", "type": "M"}):
        try:
            rr = S.post(K + f"/stat/list/{name}.es", data=data, headers={"Referer": K + "/menu.es?mid=a10101000000"}, timeout=(10, 60))
            line = f"POST {name} {data} -> {rr.status_code} {len(rr.text)} :: " + re.sub(r"\s+", " ", rr.text)[:400]
        except Exception as e:  # noqa: BLE001
            line = f"POST {name} ERR {type(e).__name__}"
        print(line, flush=True)
        log.append(line)
D = "https://www.data.go.kr/tcs/dss/selectDataSetList.do?keyword="
for kw in ("%ED%95%9C%EA%B5%AD%EA%B0%80%EC%8A%A4%EA%B3%B5%EC%82%AC+LNG", "LNG+%EB%8F%84%EC%9E%85", "%EC%A0%84%EB%A0%A5%EA%B1%B0%EB%9E%98%EC%86%8C+%EB%B0%9C%EC%A0%84%EB%9F%89"):
    t = S.get(D + kw, timeout=(10, 60)).text
    items = re.findall(r'<a[^>]+href="[^"]*/data/(\d+)/(fileData|openapi)\.do"[^>]*>(.*?)</a>', t, re.S)
    for ident, kind, title in items[:15]:
        line = f"DATASET {kw[:20]} {ident} {kind} {re.sub(r'<[^>]+>|\\s+', ' ', title).strip()[:120]}"
        print(line, flush=True)
        log.append(line)
open(os.path.join(OUT, "probe13.txt"), "w", encoding="utf-8").write("\n".join(log))
