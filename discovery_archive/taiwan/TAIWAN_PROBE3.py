"""Probe 3: Taipower 10-minute unit history JSON, WRA reservoir open data, ESIST/EA statistics API, customs."""
import re
import json
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)


def get(u, **k):
    try:
        return S.get(u, timeout=60, **k)
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


def txt(h):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", h, flags=re.S)).strip()


print("== Taipower unit history")
for n in ("d006010", "d006011", "d006009", "d006017", "d006020", "d006021", "d006022"):
    for ext in ("json", "csv"):
        u = f"https://service.taipower.com.tw/data/opendata/apply/file/{n}/001.{ext}"
        r = get(u, stream=True)
        if r is None:
            continue
        head = next(r.iter_content(1500), b"")
        print(r.status_code, u, r.headers.get("content-length"), head[:700].decode("utf-8-sig", "replace").replace("\n", " "))
        r.close()

print("== data.gov.tw")
for i in (45501,):
    r = get(f"https://data.gov.tw/api/v2/rest/dataset/{i}")
    if r is not None:
        j = r.json()["result"]
        print(j.get("title"), j.get("updateFrequency"), j.get("coverageStartedDate"), j.get("modifiedDate"), j.get("notes", "")[:300])
        for d in j.get("distribution", []):
            print("  ", d.get("resourceFormat"), d.get("resourceDownloadUrl"), d.get("resourceAmount"),
                  [f["name"] for f in d.get("resourceField", [])][:30])
for kw in ("水庫", "天然氣", "液化天然氣", "進口貨品"):
    try:
        r = S.post("https://data.gov.tw/api/front/dataset/list", json={"keyword": kw, "size": 10, "page": 1, "sort": "_score"}, timeout=40)
        j = r.json()["payload"]
        print("SEARCH", kw, j.get("search_count"))
        for x in j["search_result"][:10]:
            print("   ", x.get("nid"), x.get("agency_name"), "|", x.get("title"))
    except Exception as e:  # noqa: BLE001
        print("search err", kw, e)

print("== WRA opendata")
r = get("https://opendata.wra.gov.tw/")
if r is not None:
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S)[:150]:
        t = txt(t)
        if re.search(r"水庫|蓄水|API|api|Reservoir", t + h):
            print("   ", h[:120], "|", t[:60])
for u in ("https://opendata.wra.gov.tw/openapi/swagger/index.html", "https://opendata.wra.gov.tw/Service/OpenData.aspx?format=json&id=1602CA19-B224-4CC3-AA31-11B1B124530F",
          "https://opendata.wra.gov.tw/api/v1/Reservoir/Daily", "https://opendata.wra.gov.tw/OpenAPI/api/OpenData/50C8256D-30C5-4B8D-9B84-2E14D5C6DF71/Data",
          "https://opendata.wra.gov.tw/OpenAPI/api/OpenData/1602CA19-B224-4CC3-AA31-11B1B124530F/Data?format=json"):
    r = get(u)
    if r is not None:
        print(r.status_code, len(r.content), u, r.text[:400].replace("\n", " "))

print("== ESIST / EA statistics site")
for u in ("https://ea01.moeaea.gov.tw/a0303/02/database/api/", "https://ea01.moeaea.gov.tw/a0303/02/database/search/electric-generation/",
          "https://ea01.moeaea.gov.tw/a0303/02/newest/monthly/", "https://ea01.moeaea.gov.tw/a0303/02/database/search/energy-index/",
          "https://ea01.moeaea.gov.tw/a0303/02/en/", "https://www.esist.org.tw/database/api/", "https://www.esist.org.tw/newest/monthly/"):
    r = get(u)
    if r is None:
        continue
    print(r.status_code, len(r.content), u)
    print("   TEXT:", txt(r.text)[:1500])
    print("   API-ish:", sorted(set(re.findall(r'["\'(]((?:https?:)?/[^"\' )]*(?:api|json|xlsx?|csv|ods|download)[^"\' )]*)', r.text, re.I)))[:40])

print("== customs")
r = get("https://www.customs.gov.tw/")
if r is not None:
    print(r.status_code, r.text[:700])
for u in ("https://www.customs.gov.tw/Statistics/index", "https://web02.mof.gov.tw/njswww/WebMain.aspx?sys=100&funid=dmain&ymt=11500",
          "https://cuswebo.trade.gov.tw/FSC3010F", "https://portal.sw.nat.gov.tw/", "https://www.trade.gov.tw/", "https://cus.trade.gov.tw/"):
    r = get(u)
    if r is not None:
        print(r.status_code, len(r.content), u, txt(r.text)[:300])
