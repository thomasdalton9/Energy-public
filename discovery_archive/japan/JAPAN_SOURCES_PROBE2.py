"""Japan sources probe 2 (manual): area supply-demand CSV URL guesses per TSO, JEPX csv, MOF customs data files."""
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u, headers=None, **kw):
    h = dict(UA)
    h.update(headers or {})
    try:
        r = requests.get(u, headers=h, timeout=(8, 40), **kw)
        return r
    except Exception as e:  # noqa: BLE001
        print("  FAIL", type(e).__name__, str(e)[:100])


def head(u, headers=None, n=4):
    print("-" * 80)
    print(u)
    r = get(u, headers)
    if r is None:
        return None
    print(f"  HTTP {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
    if r.status_code == 200 and len(r.content):
        for enc in ("cp932", "utf-8"):
            try:
                t = r.content.decode(enc)
                for line in t.splitlines()[:n]:
                    print("   |", line[:400])
                print("   lines:", len(t.splitlines()))
                break
            except Exception:  # noqa: BLE001
                pass
    return r


def links(u, pat, maxn=60, headers=None):
    print("-" * 80)
    print("LINKS", u)
    r = get(u, headers)
    if r is None or r.status_code != 200:
        print("  ", None if r is None else r.status_code)
        return
    t = r.content.decode(r.apparent_encoding or "utf-8", "replace")
    ls = [l for l in dict.fromkeys(re.findall(r'(?:href|src)=["\']([^"\'#]+)', t)) if re.search(pat, l, re.I)]
    for l in ls[:maxn]:
        print("   ", l)
    for m in re.finditer(r"[^\"'\s]*csv[^\"'\s]*", t, re.I):
        pass
    snippets = [t[max(0, m.start() - 80):m.end() + 80].replace("\n", " ") for m in re.finditer(r"csv", t, re.I)][:6]
    for s in snippets:
        print("   ~", s)


# 1 HEPCO file format
head("https://www.hepco.co.jp/network/con_service/public_document/supply_demand_results/csv/eria_jukyu_202609_01.csv", n=8)
# Tokyo: guesses + referer
for u in ["https://www.tepco.co.jp/forecast/html/images/eria_jukyu_202609_03.csv",
          "https://www.tepco.co.jp/forecast/html/images/juyo-2025.csv",
          "https://www.tepco.co.jp/forecast/html/images/juyo-d1-j.csv"]:
    head(u, {"Referer": "https://www.tepco.co.jp/forecast/html/area_jukyu-j.html", "Accept-Language": "ja"})
# Tohoku
head("https://setsuden.nw.tohoku-epco.co.jp/common/demand/eria_jukyu_202609_02.csv")
links("https://setsuden.nw.tohoku-epco.co.jp/download.html", r"csv|zip|xls|download")
links("https://nw.tohoku-epco.co.jp/", r"juyo|jukyu|setsuden|csv")
# Chubu
head("https://powergrid.chuden.co.jp/denki_yoho_content_data/download_csv/eria_jukyu_202609_04.csv")
links("https://powergrid.chuden.co.jp/denkiyoho/", r"csv|jukyu|download|data")
# Hokuriku
head("https://www.rikuden.co.jp/nw/denki-yoho/csv/eria_jukyu_202609_05.csv")
links("https://www.rikuden.co.jp/nw/denki-yoho/results_jukyu.html", r"csv|jukyu|download")
links("https://www.rikuden.co.jp/nw/denki-yoho/", r"csv|jukyu|download|result")
# Kansai
head("https://www.kansai-td.co.jp/yamasou/eria_jukyu_202609_06.csv")
links("https://www.kansai-td.co.jp/denkiyoho/download/", r"csv|jukyu|download|zip|xls")
# Chugoku
links("https://www.energia.co.jp/nw/jukyuu/eria_jukyu.html", r".")
head("https://www.energia.co.jp/nw/jukyuu/sys/eria_jukyu_202609_07.csv")
# Shikoku
links("https://www.yonden.co.jp/nw/supply_demand/data_download.html", r".")
head("https://www.yonden.co.jp/nw/supply_demand/csv/eria_jukyu_202609_08.csv")
# Kyushu
head("https://www.kyuden.co.jp/td_power_usages/csv/eria_jukyu_202609_09.csv")
links("https://www.kyuden.co.jp/td_power_usages/pc.html", r"csv|jukyu|download|zip|xls")
links("https://www.kyuden.co.jp/td/service/wheeling/disclosure.html", r"jukyu|eria|supply", 30)
# Okinawa
links("https://www.okinawa-epco.co.jp/", r"jukyu|csv|yoho")
# JEPX
links("https://www.jepx.jp/electricpower/market-data/spot/", r"csv|spot|download|js/")
head("https://www.jepx.jp/js/csv_read.php?dir=spot_summary&file=spot_summary_2025.csv",
     {"Referer": "https://www.jepx.jp/electricpower/market-data/spot/"})
head("https://www.jepx.jp/js/csv_read.php?dir=spot_summary&file=spot_summary_2025.csv",
     {"Referer": "https://www.jepx.jp/electricpower/market-data/spot/", "Accept-Language": "ja", "Accept": "text/csv,*/*"})
head("https://www.jepx.jp/electricpower/market-data/spot/ave_day.html")
# MOF customs
links("https://www.customs.go.jp/toukei/info/index_e.htm", r"\.htm|\.csv|\.xls|\.zip", 60)
links("https://www.customs.go.jp/toukei/srch/indexe.htm", r"\.htm|\.csv|\.xls|\.zip", 70)
links("https://www.customs.go.jp/toukei/suii/html/time_e.htm", r"\.htm|\.csv|\.xls|\.zip", 70)
links("https://www.customs.go.jp/toukei/latest/index.htm", r"\.csv|\.xls|\.zip|e\.htm", 60)
r = get("https://www.customs.go.jp/toukei/shinbun/trade-st/2026/2026085.pdf")
if r is not None and r.status_code == 200:
    import pymupdf
    d = pymupdf.open(stream=r.content, filetype="pdf")
    print("PDF pages", len(d))
    for i, pg in enumerate(d):
        t = pg.get_text()
        if "液化天然ガス" in t or "LNG" in t:
            print("page", i)
            idx = t.find("液化天然ガス")
            print(t[max(0, idx - 300): idx + 600])
            break
