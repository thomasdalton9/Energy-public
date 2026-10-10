"""Probe 2: data.gov.tw resources, MOEA Energy Administration menus and file links, WRA API, esist, customs."""
import re
import sys
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)


def get(u, **k):
    try:
        return S.get(u, timeout=40, **k)
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


print("== data.gov.tw datasets")
for i in (19995, 37331):
    r = get(f"https://data.gov.tw/api/v2/rest/dataset/{i}")
    if r is not None:
        print(i, r.text[:3000])
for body in ({"keyword": "發購電量", "size": 20, "page": 1}, {"keyword": "水庫", "size": 20, "page": 1}):
    try:
        r = S.post("https://data.gov.tw/api/front/dataset/list", json=body, timeout=40)
        print("POST", body, r.status_code, r.text[:1500])
    except Exception as e:  # noqa: BLE001
        print("POST err", e)

print("== moeaea menus")
base = "https://www.moeaea.gov.tw/ECW/populace/content/SubMenu.aspx?menu_id="
for m in range(1500, 1600):
    r = get(base + str(m))
    if r is None or r.status_code != 200:
        continue
    t = re.search(r"<title>(.*?)</title>", r.text, re.S)
    files = re.findall(r'href="([^"]+\.(?:xls|xlsx|ods|csv|pdf)[^"]*)"', r.text, re.I)
    print(m, (t.group(1).strip()[:80] if t else ""), "files:", len(files), files[:4])
for m in (1540, 1541, 1542):
    r = get(base + str(m))
    if r is not None:
        links = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S)
        for h, tx in links:
            tx = re.sub(r"<.*?>|\s+", " ", tx).strip()
            if re.search(r"月報|Monthly|天然氣|LNG|電力|統計|xls|ods|csv", tx + h, re.I):
                print("   ", m, h[:140], "|", tx[:70])

print("== WRA")
r = get("https://fhy.wra.gov.tw/ReservoirPage_2011/Statistics.aspx")
if r is not None:
    print(r.text[:1700])
    for js in re.findall(r'src="([^"]+\.js[^"]*)"', r.text)[:6]:
        print("js", js)
for u in ("https://fhy.wra.gov.tw/Api/v2/Reservoir/Station", "https://fhy.wra.gov.tw/WraApi/v1/Reservoir/Station",
          "https://fhy.wra.gov.tw/WraApi/swagger/index.html", "https://opendata.wra.gov.tw/",
          "https://www.wra.gov.tw/cp.aspx?n=49961", "https://data.gov.tw/dataset/45501", "https://data.gov.tw/api/v2/rest/dataset/45501",
          "https://data.gov.tw/api/v2/rest/dataset/25256", "https://data.gov.tw/api/v2/rest/dataset/14419"):
    r = get(u)
    if r is not None:
        print(r.status_code, len(r.content), u, r.text[:400].replace("\n", " "))

print("== esist")
r = get("https://www.esist.org.tw/Database/DatabaseHome")
if r is not None:
    for js in sorted(set(re.findall(r'(?:src|href)="([^"]+)"', r.text)))[:60]:
        print("   ", js)
for u in ("https://www.esist.org.tw/api/", "https://www.esist.org.tw/Database/Download"):
    r = get(u)
    if r is not None:
        print(r.status_code, u, r.text[:200].replace("\n", " "))

print("== customs / mof / cpc / others")
for u in ("https://portal.sw.nat.gov.tw/APGA/GA30E", "https://www.customs.gov.tw/", "https://web02.mof.gov.tw/njswww/WebMain.aspx?sys=210&funid=defjsptgl",
          "https://service.mof.gov.tw/", "https://www.mof.gov.tw/", "https://www.cpc.com.tw/en/cp.aspx?n=2362",
          "https://www.moeaea.gov.tw/ECW/populace/content/SubMenu.aspx?menu_id=1544",
          "https://www.moeaea.gov.tw/ECW/populace/content/wHandMenuFile.ashx?menu_id=1540",
          "https://www.moeaea.gov.tw/ECW/english/content/ContentLink.aspx?menu_id=1540",
          "https://www.taipower.com.tw/",
          "https://service.taipower.com.tw/data/opendata/apply/file/d006005/001.csv"):
    r = get(u)
    if r is not None:
        print(r.status_code, len(r.content), u, r.text[:250].replace("\n", " "))
        if "d006005" in u:
            lines = r.content.decode("utf-8-sig", "replace").splitlines()
            print(len(lines)); print(lines[0]); print(lines[1]); print(lines[-1])
