"""Probe 10: WRA reservoir names - warning-facility dataset fields, swagger initialiser, dataset ids around."""
import json
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
D = "discovery_archive/results/taiwan"
LOG = []


def get(u, **k):
    try:
        return S.get(u, timeout=90, **k)
    except Exception as e:  # noqa: BLE001
        LOG.append(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


W = "https://opendata.wra.gov.tw"
r = get(f"{W}/api/v2/6c82a943-5ef9-4c31-9854-fb08c376bb3f?format=JSON&limit=4")
LOG.append(f"WARN {r.status_code if r is not None else None} {r.text[:900] if r is not None else ''}")
for u in ("/openapi/swagger/js/swagger-initializer.js", "/openapi/swagger/index.html", "/openapi/v1/swagger.json",
          "/openapi/swagger/doc.json", "/openapi/docs/openapi.json", "/openapi/openapi.json", "/api/v1/datasets?topic_name=%E6%B0%B4%E5%BA%AB%E8%88%87%E5%A0%B0%E5%A3%A9&page=1",
          "/api/v1/datasets?query=%E6%B0%B4%E5%BA%AB%E6%AF%8F%E6%97%A5&page=1", "/api/v1/topics"):
    r = get(W + u)
    if r is not None:
        LOG.append(f"TRY {r.status_code} {len(r.content)} {u} {r.text[:500].replace(chr(10), ' ')}")
hits = []
for i in list(range(45300, 45380)) + list(range(45620, 45700)):
    r = get(f"https://data.gov.tw/api/v2/rest/dataset/{i}")
    try:
        j = r.json()["result"]
        if "opendata.wra" in json.dumps(j.get("distribution", [])) or "水利署" in j.get("title", ""):
            hits.append((i, j["title"]))
    except Exception:  # noqa: BLE001
        pass
LOG.append("DG " + json.dumps(hits, ensure_ascii=False))
open(D + "/probe10_log.txt", "w", encoding="utf-8").write("\n".join(LOG))
print(LOG[0][:300])
