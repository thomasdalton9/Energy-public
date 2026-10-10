"""Japan sources probe 6 (manual): how the Chubu and Kansai pages build their monthly area supply-demand CSV links."""
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u):
    try:
        return requests.get(u, headers=UA, timeout=(8, 40))
    except Exception as e:  # noqa: BLE001
        print("FAIL", u, type(e).__name__)


def show(u, pat, width=220, maxn=12):
    r = get(u)
    if r is None:
        return
    t = r.content.decode("utf-8", "replace")
    print("====", u, r.status_code, len(t))
    for m in list(re.finditer(pat, t))[:maxn]:
        print("  ~", t[max(0, m.start() - width): m.end() + width].replace("\n", " "))
    return t


# Chubu
t = show("https://powergrid.chuden.co.jp/denkiyoho/index.html", r"eria|month_grid|supply_demand|\.csv", 160)
for js in ("https://powergrid.chuden.co.jp/denkiyoho/resource/js/create-download-link.js",
           "https://powergrid.chuden.co.jp/denkiyoho/resource/js/get-data.js"):
    show(js, r"csv|url|http|\.\./", 150, 14)
# Kansai
show("https://www.kansai-td.co.jp/denkiyoho/download/js/jisseki.js", r"csv|url|http|fetch|axios|\.json", 150, 14)
show("https://www.kansai-td.co.jp/denkiyoho/area-performance/index.html", r"csv|jisseki|download", 150, 10)
