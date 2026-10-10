"""Probe Taiwan sources from GitHub Actions: status, type, size and the head of each response."""
import sys
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
URLS = [
 "https://service.taipower.com.tw/data/opendata/apply/file/d006001/001.json",
 "https://www.taipower.com.tw/d006/loadGraph/loadGraph/data/genary.json",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006002/001.json",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006003/001.json",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006004/001.json",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006005/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006006/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006007/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006008/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006009/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006011/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006012/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006013/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006014/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006015/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006016/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006017/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006018/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006019/001.csv",
 "https://service.taipower.com.tw/data/opendata/apply/file/d006020/001.csv",
 "https://data.gov.tw/api/front/dataset/list?keyword=%E5%8F%B0%E7%81%A3%E9%9B%BB%E5%8A%9B%E5%85%AC%E5%8F%B8&size=30",
 "https://data.gov.tw/api/front/dataset/list?keyword=%E7%99%BC%E8%B3%BC%E9%9B%BB%E9%87%8F&size=30",
 "https://data.gov.tw/dataset/19995",
 "https://data.gov.tw/api/v2/rest/dataset/19995",
 "https://data.gov.tw/api/v2/rest/dataset/37331",
 "https://www.taipower.com.tw/tc/page.aspx?mid=96",
 "https://www.taipower.com.tw/en/page.aspx?mid=4492",
 "https://www.moeaea.gov.tw/ECW/populace/content/SubMenu.aspx?menu_id=1540",
 "https://www.moeaea.gov.tw/ECW/populace/content/ContentLink.aspx?menu_id=1540",
 "https://www.moeaea.gov.tw/ecw/english/content/SubMenu.aspx?menu_id=1540",
 "https://www.esist.org.tw/Database/DatabaseHome",
 "https://www.esist.org.tw/",
 "https://web02.mof.gov.tw/njswww/WebMain.aspx?sys=100&funid=dmain",
 "https://portal.sw.nat.gov.tw/APGA/GA30",
 "https://data.gov.tw/api/front/dataset/list?keyword=%E5%A4%A9%E7%84%B6%E6%B0%A3&size=30",
 "https://data.gov.tw/api/front/dataset/list?keyword=%E8%83%BD%E6%BA%90%E7%B5%B1%E8%A8%88&size=30",
 "https://data.wra.gov.tw/Service/OpenData.aspx?format=json&id=1602CA19-B224-4CC3-AA31-11B1B124530F",
 "https://data.wra.gov.tw/Service/OpenData.aspx?format=json&id=50C8256D-30C5-4B8D-9B84-2E14D5C6DF71",
 "https://fhy.wra.gov.tw/WraApi/v1/Reservoir/Info?$top=3&$format=JSON",
 "https://fhy.wra.gov.tw/WraApi/v1/Reservoir/Daily?$top=3&$format=JSON",
 "https://fhy.wra.gov.tw/WraApi/v1/Reservoir/Daily?$top=3&$format=JSON&$filter=ObservationTime ge 2020-01-01",
 "https://fhy.wra.gov.tw/ReservoirPage_2011/Statistics.aspx",
 "https://www.wra.gov.tw/",
 "https://opendata.cwa.gov.tw/index",
 "https://www.moea.gov.tw/",
 "https://www.cpc.com.tw/",
 "https://www.cpc.com.tw/en/",
 "https://www.moeaea.gov.tw/ECW/populace/news/News.aspx?kind=1&menu_id=40",
 "https://www.moeaea.gov.tw/ECW/populace/home/Home.aspx",
 "https://www.moeaea.gov.tw/ECW/english/home/English.aspx",
 "https://www.dgbas.gov.tw/",
 "https://nstatdb.dgbas.gov.tw/dgbasall/webMain.aspx?sys=210&funid=dgmain",
 "https://www.imf.org/external/datamapper/api/v1/NGDP_RPCH/TWN",
 "https://api.worldbank.org/v2/country/TWN?format=json",
 "https://api.worldbank.org/v2/country/CHN/indicator/NY.GDP.MKTP.CD?format=json&per_page=2",
]
for u in (sys.argv[1:] or URLS):
    try:
        r = requests.get(u, headers=UA, timeout=40)
        t = r.content[:400].decode("utf-8", "replace").replace("\n", " ")
        print(f"[{r.status_code}] {len(r.content):>9} {r.headers.get('content-type','')[:30]:30} {u}\n      {t}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"[ERR] {u}\n      {type(e).__name__}: {str(e)[:150]}", flush=True)
