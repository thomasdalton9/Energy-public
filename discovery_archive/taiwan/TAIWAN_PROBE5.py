"""Probe 5 (compact): ESIST monthly tables, WRA reservoir API depth and reservoir names."""
import json
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
OUT = []


def p(*a):
    OUT.append(" ".join(str(x) for x in a).replace("\n", " ")[:330])


def get(u, **k):
    try:
        return S.get(u, timeout=90, **k)
    except Exception as e:  # noqa: BLE001
        p("ERR", u, type(e).__name__, str(e)[:80])


E = "https://ea01.moeaea.gov.tw/a0303/02/api"
r = get("https://ea01.moeaea.gov.tw/a0303/02/api/pages/database/api")
paths = json.loads(r.text)["paths"] if r is not None else {}
p("PATHS", " | ".join(f"{k.replace('/monthly/', '')}={v['get']['summary'][5:].strip()[:14]}" for k, v in paths.items()))
for t in ("3/2", "3/3", "2/2", "3/1", "3/4", "4/1", "4/2"):
    r = get(f"{E}/v1/zone/monthly/{t}")
    if r is None:
        continue
    p("T", t, r.status_code, len(r.content), r.headers.get("content-type"), r.text[:260])
    try:
        j = r.json()
        def shape(o, d=0):
            if isinstance(o, dict):
                return "{" + ",".join(f"{k}:{shape(v, d + 1) if d < 2 else '..'}" for k, v in list(o.items())[:6]) + "}"
            if isinstance(o, list):
                return f"[{len(o)}x " + (shape(o[0], d + 1) if o and d < 3 else "") + "]"
            return type(o).__name__
        p("   shape", shape(j))
    except Exception as e:  # noqa: BLE001
        p("   notjson", e)

W = "https://opendata.wra.gov.tw/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679"
for q in ("limit=1000&offset=0", "limit=5&offset=100000", "limit=5&offset=1000000", "limit=5&offset=5000000"):
    r = get(f"{W}?format=JSON&{q}")
    if r is not None:
        try:
            j = r.json()
            obs = [x.get("observationtime") for x in j]
            p("WRA", q, r.status_code, len(j), "obs", min(obs) if obs else None, max(obs) if obs else None,
              "ids", len({x.get("reservoiridentifier") for x in j}))
        except Exception:  # noqa: BLE001
            p("WRA", q, r.status_code, r.text[:200])
for q in ("filter=reservoiridentifier%20eq%2010501&limit=3", "reservoiridentifier=10501&limit=3", "observationtime=2024-01-01&limit=3"):
    r = get(f"{W}?format=JSON&{q}")
    if r is not None:
        p("WRAF", q, r.status_code, r.text[:300])
for u in ("https://opendata.wra.gov.tw/openapi/swagger/v1/swagger.json", "https://opendata.wra.gov.tw/openapi/swagger.json",
          "https://opendata.wra.gov.tw/api/v2/datasets?keyword=%E6%B0%B4%E5%BA%AB", "https://opendata.wra.gov.tw/api/datasets?keyword=%E6%B0%B4%E5%BA%AB"):
    r = get(u)
    if r is not None:
        p("WRAX", r.status_code, len(r.content), u, r.text[:200])
for i in list(range(45490, 45512)) + [9186, 9187, 25251, 25252, 25253]:
    r = get(f"https://data.gov.tw/api/v2/rest/dataset/{i}")
    if r is not None:
        try:
            j = r.json()["result"]
            if re.search("水庫|水情|蓄水", j["title"]):
                p("DG", i, j["title"], [d.get("resourceDownloadUrl") for d in j["distribution"]][:2], j.get("coverageStartedDate"))
        except Exception:  # noqa: BLE001
            pass
print("\n".join(OUT))
