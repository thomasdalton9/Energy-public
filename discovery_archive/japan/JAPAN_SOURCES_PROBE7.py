"""Japan sources probe 7 (manual): guessed Chubu / Kansai area supply-demand CSV locations; compact JS grep."""
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u):
    try:
        return requests.get(u, headers=UA, timeout=(8, 40))
    except Exception as e:  # noqa: BLE001
        print("FAIL", u, type(e).__name__)


def st(u, n=2):
    r = get(u)
    if r is None:
        return
    s = f"{r.status_code} {len(r.content)}b {r.headers.get('content-type')}"
    print(u, "->", s)
    if r.status_code == 200 and "html" not in (r.headers.get("content-type") or ""):
        for enc in ("cp932", "utf-8"):
            try:
                for line in r.content.decode(enc).splitlines()[:n]:
                    print("    |", line[:220])
                break
            except Exception:  # noqa: BLE001
                pass


def grep(u, pat, maxn=20):
    r = get(u)
    if r is None:
        return
    t = r.content.decode("utf-8", "replace")
    print("GREP", u, r.status_code, len(t))
    for ln in t.splitlines():
        if re.search(pat, ln) and len(ln) < 400:
            print("   ", ln.strip()[:300])
            maxn -= 1
            if maxn <= 0:
                break


K = "https://www.kansai-td.co.jp/interchange/denkiyoho/area-performance/"
for u in (K + "eria_jukyu_202608_06.csv", K + "jisseki.json", K + "list.json",
          "https://www.kansai-td.co.jp/interchange/denkiyoho/download/eria_jukyu_202608_06.csv",
          "https://www.kansai-td.co.jp/interchange/denkiyoho/area-performance/csv/eria_jukyu_202608_06.csv",
          "https://www.kansai-td.co.jp/interchange/denkiyoho/download/jisseki_list.json",
          "https://www.kansai-td.co.jp/interchange/denkiyoho/download/list.json"):
    st(u)
grep("https://www.kansai-td.co.jp/denkiyoho/download/js/jisseki.js", r"json|csv|path|name|label|\.get\(")
grep("https://www.kansai-td.co.jp/denkiyoho/download/", r"json|csv|jukyu|data-")
C = "https://powergrid.chuden.co.jp/denki_yoho_content_data/"
for u in (C + "eria_jukyu_202608_04.csv", C + "download_csv/eria_jukyu_202608_04.csv", C + "eriajukyu_data/eria_jukyu_202608_04.csv",
          "https://powergrid.chuden.co.jp/denkiyoho/eriajukyu_data/csv/eria_jukyu_202608_04.csv",
          "https://powergrid.chuden.co.jp/denki_yoho_content_data/past/eria_jukyu_202608_04.csv",
          "https://powergrid.chuden.co.jp/denki_yoho_content_data/month/eria_jukyu_202608_04.csv",
          "https://powergrid.chuden.co.jp/denkiyoho/index.html"):
    st(u)
grep("https://powergrid.chuden.co.jp/denkiyoho/index.html", r"eria|month_grid|jukyu_|_link")
grep("https://powergrid.chuden.co.jp/denkiyoho/resource/js/create-download-link.js", r"csv|url|href|path|http")
grep("https://powergrid.chuden.co.jp/denkiyoho/resource/js/get-data.js", r"csv|url|href|path|http", 12)
grep("https://powergrid.chuden.co.jp/denkiyoho/eriajukyu_data/", r"script|json|csv", 14)
