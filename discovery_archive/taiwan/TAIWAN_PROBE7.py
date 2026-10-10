"""Probe 7: save raw responses under discovery_archive/results/taiwan/ (committed by the workflow) so they can be read
locally: every ESIST monthly table, the WRA reservoir snapshot, heads of the Taipower files; hunt the WRA reservoir
name list in the open-data site's scripts."""
import json
import os
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
D = "discovery_archive/results/taiwan"
os.makedirs(D + "/esist", exist_ok=True)
LOG = []


def get(u, **k):
    try:
        return S.get(u, timeout=120, **k)
    except Exception as e:  # noqa: BLE001
        LOG.append(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


E = "https://ea01.moeaea.gov.tw/a0303/02/api"
r = get(E + "/pages/database/api")
paths = list(json.loads(r.text)["paths"])
for pth in paths:
    if not pth.startswith("/monthly/"):
        continue
    rr = get(E + "/v1/zone" + pth)
    if rr is not None and rr.status_code == 200:
        open(f"{D}/esist/{pth[9:].replace('/', '_')}.json", "w", encoding="utf-8").write(rr.text)
        LOG.append(f"esist {pth} {len(rr.content)}")
    else:
        LOG.append(f"esist {pth} FAIL {getattr(rr, 'status_code', None)}")
rr = get("https://opendata.wra.gov.tw/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679?format=JSON&limit=1000")
open(D + "/wra_snapshot.json", "w", encoding="utf-8").write(rr.text)
rr = get("https://service.taipower.com.tw/data/opendata/apply/file/d006005/001.csv")
open(D + "/taipower_d006005_head.csv", "w", encoding="utf-8").write("\n".join(rr.content.decode("utf-8-sig").splitlines()[:6]))
rr = get("https://service.taipower.com.tw/data/opendata/apply/file/d006001/001.json")
open(D + "/taipower_d006001_live.json", "w", encoding="utf-8").write(rr.content.decode("utf-8-sig"))
rr = S.get("https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.json", stream=True, timeout=120)
open(D + "/taipower_d006010_head.json", "w", encoding="utf-8").write(next(rr.iter_content(300000)).decode("utf-8", "ignore"))
rr.close()
# hunt WRA reservoir names
h = get("https://opendata.wra.gov.tw/datasets?topic_name=%E6%B0%B4%E5%BA%AB%E8%88%87%E5%A0%B0%E5%A3%A9&page=1").text
for js in re.findall(r'(?:src|href)="(/_nuxt/[^"]+\.js)"', h):
    t = get("https://opendata.wra.gov.tw" + js)
    if t is None:
        continue
    hits = sorted(set(re.findall(r'[\w/\-.:]*api/v\d[\w/\-.?=&{}$]*', t.text)))[:8]
    LOG.append(f"js {js} {len(t.text)} {hits}")
for q in ("%E6%B0%B4%E5%BA%AB%E6%AF%8F%E6%97%A5%E7%87%9F%E9%81%8B%E7%8B%80%E6%B3%81", "ReservoirName", "reservoir"):
    t = get(f"https://opendata.wra.gov.tw/api/v2/search?keyword={q}")
    LOG.append(f"search {q} {getattr(t, 'status_code', None)} {(t.text[:200] if t is not None else '')}")
open(D + "/probe7_log.txt", "w", encoding="utf-8").write("\n".join(LOG))
print("\n".join(l[:200] for l in LOG[-12:]))
