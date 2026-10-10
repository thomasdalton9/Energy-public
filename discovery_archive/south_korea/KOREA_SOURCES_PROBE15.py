"""
Probe 15 (manual workflow south_korea_probe.yml): data.go.kr dataset search (HTML list pages, keyless) for gas / oil /
power file datasets of Korean agencies; prints id, type, title and the list page's own modification text; saved to
discovery_archive/results/south_korea/probe15.txt.
"""
import os
import re
from urllib.parse import quote

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers.update(H)
log = []
KW = ["한국가스공사 판매량", "한국가스공사 도입", "천연가스 수입", "한국가스공사 LNG 재고", "LNG 인수기지", "용도별 천연가스 판매",
      "도시가스 판매량", "한국석유공사 원유 수입", "석유제품 재고", "원유 도입", "전력거래소 월별 발전량", "발전원별 월별 발전량",
      "한국가스공사 월별", "천연가스 수급", "에너지통계 월보", "한국전력 판매전력량", "LNG 수입량"]
seen = set()
for kw in KW:
    url = "https://www.data.go.kr/tcs/dss/selectDataSetList.do?keyword=" + quote(kw)
    try:
        t = S.get(url, timeout=(10, 60)).text
    except Exception as e:  # noqa: BLE001
        log.append(f"ERR {kw} {type(e).__name__}")
        continue
    items = re.findall(r'<a[^>]+href="[^"]*/data/(\d+)/(fileData|openapi)\.do"[^>]*>(.*?)</a>', t, re.S)
    print(f"== {kw}: {len(items)}", flush=True)
    log.append(f"== {kw}: {len(items)}")
    for ident, kind, title in items[:25]:
        if (ident, kind) in seen:
            continue
        seen.add((ident, kind))
        clean = re.sub(r"<[^>]+>|\s+", " ", title).strip()[:110]
        line = f"{ident} {kind} {clean}"
        print(line, flush=True)
        log.append(line)
open(os.path.join(OUT, "probe15.txt"), "w", encoding="utf-8").write("\n".join(log))
